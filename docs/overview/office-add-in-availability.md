---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos com suporte para o Excel, OneNote, Outlook, PowerPoint, Project e Word.
ms.date: 08/13/2019
localization_priority: Priority
ms.openlocfilehash: a3c580f32ad7cd384309a9b53e55ea488a470a90
ms.sourcegitcommit: f781d7cfd980cd866d6d1d00c5b9d16c8a4b7f9b
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/20/2019
ms.locfileid: "37053323"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="715a0-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="715a0-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="715a0-104">Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="715a0-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="715a0-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que são compatíveis atualmente com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="715a0-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="715a0-106">A versão inicial do Office 2016 instalada por meio do MSI apenas contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="715a0-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="715a0-107">Para saber mais sobre o histórico de atualizações de várias versões do Office, confira a seção[Confira também](#see-also).</span><span class="sxs-lookup"><span data-stu-id="715a0-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="715a0-108">Excel</span><span class="sxs-lookup"><span data-stu-id="715a0-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="715a0-109">Plataforma</span><span class="sxs-lookup"><span data-stu-id="715a0-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="715a0-110">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="715a0-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="715a0-111">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="715a0-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="715a0-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="715a0-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-113">Office na Web</span><span class="sxs-lookup"><span data-stu-id="715a0-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="715a0-114">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="715a0-114">- TaskPane</span></span><br><span data-ttu-id="715a0-115">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="715a0-115">
        - Content</span></span><br><span data-ttu-id="715a0-116">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="715a0-116">
        - Custom Functions</span></span><br><span data-ttu-id="715a0-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="715a0-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="715a0-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="715a0-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="715a0-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="715a0-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="715a0-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="715a0-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="715a0-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="715a0-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="715a0-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="715a0-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="715a0-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="715a0-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="715a0-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="715a0-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="715a0-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="715a0-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="715a0-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="715a0-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="715a0-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-128">
        - BindingEvents</span></span><br><span data-ttu-id="715a0-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="715a0-129">
        - CompressedFile</span></span><br><span data-ttu-id="715a0-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-130">
        - DocumentEvents</span></span><br><span data-ttu-id="715a0-131">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="715a0-131">
        - File</span></span><br><span data-ttu-id="715a0-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-132">
        - MatrixBindings</span></span><br><span data-ttu-id="715a0-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-133">
        - MatrixCoercion</span></span><br><span data-ttu-id="715a0-134">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="715a0-134">
        - Selection</span></span><br><span data-ttu-id="715a0-135">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="715a0-135">
        - Settings</span></span><br><span data-ttu-id="715a0-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-136">
        - TableBindings</span></span><br><span data-ttu-id="715a0-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-137">
        - TableCoercion</span></span><br><span data-ttu-id="715a0-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-138">
        - TextBindings</span></span><br><span data-ttu-id="715a0-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-139">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-140">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="715a0-140">Office on Windows</span></span><br><span data-ttu-id="715a0-141">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="715a0-141">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="715a0-142">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="715a0-142">- TaskPane</span></span><br><span data-ttu-id="715a0-143">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="715a0-143">
        - Content</span></span><br><span data-ttu-id="715a0-144">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="715a0-144">
        - Custom Functions</span></span><br><span data-ttu-id="715a0-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="715a0-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="715a0-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="715a0-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="715a0-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="715a0-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="715a0-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="715a0-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="715a0-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="715a0-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="715a0-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="715a0-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="715a0-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="715a0-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="715a0-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="715a0-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="715a0-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="715a0-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="715a0-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="715a0-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="715a0-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="715a0-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="715a0-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="715a0-158">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-158">
        - BindingEvents</span></span><br><span data-ttu-id="715a0-159">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="715a0-159">
        - CompressedFile</span></span><br><span data-ttu-id="715a0-160">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-160">
        - DocumentEvents</span></span><br><span data-ttu-id="715a0-161">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="715a0-161">
        - File</span></span><br><span data-ttu-id="715a0-162">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-162">
        - MatrixBindings</span></span><br><span data-ttu-id="715a0-163">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-163">
        - MatrixCoercion</span></span><br><span data-ttu-id="715a0-164">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="715a0-164">
        - Selection</span></span><br><span data-ttu-id="715a0-165">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="715a0-165">
        - Settings</span></span><br><span data-ttu-id="715a0-166">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-166">
        - TableBindings</span></span><br><span data-ttu-id="715a0-167">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-167">
        - TableCoercion</span></span><br><span data-ttu-id="715a0-168">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-168">
        - TextBindings</span></span><br><span data-ttu-id="715a0-169">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-169">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-170">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="715a0-170">Office 2019 on Windows</span></span><br><span data-ttu-id="715a0-171">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="715a0-171">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="715a0-172">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="715a0-172">- TaskPane</span></span><br><span data-ttu-id="715a0-173">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="715a0-173">
        - Content</span></span><br><span data-ttu-id="715a0-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="715a0-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="715a0-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="715a0-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="715a0-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="715a0-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="715a0-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="715a0-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="715a0-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="715a0-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="715a0-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="715a0-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="715a0-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="715a0-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="715a0-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="715a0-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="715a0-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="715a0-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="715a0-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="715a0-185">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-185">- BindingEvents</span></span><br><span data-ttu-id="715a0-186">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="715a0-186">
        - CompressedFile</span></span><br><span data-ttu-id="715a0-187">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-187">
        - DocumentEvents</span></span><br><span data-ttu-id="715a0-188">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="715a0-188">
        - File</span></span><br><span data-ttu-id="715a0-189">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-189">
        - MatrixBindings</span></span><br><span data-ttu-id="715a0-190">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-190">
        - MatrixCoercion</span></span><br><span data-ttu-id="715a0-191">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="715a0-191">
        - Selection</span></span><br><span data-ttu-id="715a0-192">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="715a0-192">
        - Settings</span></span><br><span data-ttu-id="715a0-193">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-193">
        - TableBindings</span></span><br><span data-ttu-id="715a0-194">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-194">
        - TableCoercion</span></span><br><span data-ttu-id="715a0-195">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-195">
        - TextBindings</span></span><br><span data-ttu-id="715a0-196">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-196">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-197">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="715a0-197">Office 2016 on Windows</span></span><br><span data-ttu-id="715a0-198">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="715a0-198">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="715a0-199">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="715a0-199">- TaskPane</span></span><br><span data-ttu-id="715a0-200">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="715a0-200">
        - Content</span></span></td>
    <td><span data-ttu-id="715a0-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="715a0-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="715a0-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="715a0-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="715a0-204">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-204">- BindingEvents</span></span><br><span data-ttu-id="715a0-205">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="715a0-205">
        - CompressedFile</span></span><br><span data-ttu-id="715a0-206">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-206">
        - DocumentEvents</span></span><br><span data-ttu-id="715a0-207">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="715a0-207">
        - File</span></span><br><span data-ttu-id="715a0-208">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-208">
        - MatrixBindings</span></span><br><span data-ttu-id="715a0-209">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-209">
        - MatrixCoercion</span></span><br><span data-ttu-id="715a0-210">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="715a0-210">
        - Selection</span></span><br><span data-ttu-id="715a0-211">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="715a0-211">
        - Settings</span></span><br><span data-ttu-id="715a0-212">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-212">
        - TableBindings</span></span><br><span data-ttu-id="715a0-213">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-213">
        - TableCoercion</span></span><br><span data-ttu-id="715a0-214">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-214">
        - TextBindings</span></span><br><span data-ttu-id="715a0-215">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-215">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-216">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="715a0-216">Office 2013 on Windows</span></span><br><span data-ttu-id="715a0-217">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="715a0-217">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="715a0-218">
        - TaskPane</span><span class="sxs-lookup"><span data-stu-id="715a0-218">
        - TaskPane</span></span><br><span data-ttu-id="715a0-219">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="715a0-219">
        - Content</span></span></td>
    <td>  <span data-ttu-id="715a0-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="715a0-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="715a0-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="715a0-222">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-222">
        - BindingEvents</span></span><br><span data-ttu-id="715a0-223">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="715a0-223">
        - CompressedFile</span></span><br><span data-ttu-id="715a0-224">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-224">
        - DocumentEvents</span></span><br><span data-ttu-id="715a0-225">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="715a0-225">
        - File</span></span><br><span data-ttu-id="715a0-226">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-226">
        - MatrixBindings</span></span><br><span data-ttu-id="715a0-227">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-227">
        - MatrixCoercion</span></span><br><span data-ttu-id="715a0-228">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="715a0-228">
        - Selection</span></span><br><span data-ttu-id="715a0-229">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="715a0-229">
        - Settings</span></span><br><span data-ttu-id="715a0-230">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-230">
        - TableBindings</span></span><br><span data-ttu-id="715a0-231">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-231">
        - TableCoercion</span></span><br><span data-ttu-id="715a0-232">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-232">
        - TextBindings</span></span><br><span data-ttu-id="715a0-233">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-233">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-234">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="715a0-234">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="715a0-235">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="715a0-235">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="715a0-236">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="715a0-236">- TaskPane</span></span><br><span data-ttu-id="715a0-237">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="715a0-237">
        - Content</span></span></td>
    <td><span data-ttu-id="715a0-238">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-238">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="715a0-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="715a0-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="715a0-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="715a0-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="715a0-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="715a0-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="715a0-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="715a0-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="715a0-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="715a0-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="715a0-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="715a0-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="715a0-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="715a0-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="715a0-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="715a0-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="715a0-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="715a0-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="715a0-249">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-249">- BindingEvents</span></span><br><span data-ttu-id="715a0-250">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-250">
        - DocumentEvents</span></span><br><span data-ttu-id="715a0-251">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="715a0-251">
        - File</span></span><br><span data-ttu-id="715a0-252">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-252">
        - MatrixBindings</span></span><br><span data-ttu-id="715a0-253">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-253">
        - MatrixCoercion</span></span><br><span data-ttu-id="715a0-254">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="715a0-254">
        - Selection</span></span><br><span data-ttu-id="715a0-255">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="715a0-255">
        - Settings</span></span><br><span data-ttu-id="715a0-256">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-256">
        - TableBindings</span></span><br><span data-ttu-id="715a0-257">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-257">
        - TableCoercion</span></span><br><span data-ttu-id="715a0-258">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-258">
        - TextBindings</span></span><br><span data-ttu-id="715a0-259">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-259">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-260">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="715a0-260">Office apps on Mac</span></span><br><span data-ttu-id="715a0-261">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="715a0-261">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="715a0-262">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="715a0-262">- TaskPane</span></span><br><span data-ttu-id="715a0-263">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="715a0-263">
        - Content</span></span><br><span data-ttu-id="715a0-264">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="715a0-264">
        - Custom Functions</span></span><br><span data-ttu-id="715a0-265">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="715a0-265">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="715a0-266">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-266">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="715a0-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="715a0-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="715a0-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="715a0-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="715a0-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="715a0-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="715a0-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="715a0-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="715a0-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="715a0-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="715a0-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="715a0-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="715a0-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="715a0-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="715a0-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="715a0-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="715a0-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="715a0-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="715a0-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="715a0-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="715a0-278">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-278">- BindingEvents</span></span><br><span data-ttu-id="715a0-279">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="715a0-279">
        - CompressedFile</span></span><br><span data-ttu-id="715a0-280">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-280">
        - DocumentEvents</span></span><br><span data-ttu-id="715a0-281">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="715a0-281">
        - File</span></span><br><span data-ttu-id="715a0-282">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-282">
        - MatrixBindings</span></span><br><span data-ttu-id="715a0-283">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-283">
        - MatrixCoercion</span></span><br><span data-ttu-id="715a0-284">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="715a0-284">
        - PdfFile</span></span><br><span data-ttu-id="715a0-285">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="715a0-285">
        - Selection</span></span><br><span data-ttu-id="715a0-286">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="715a0-286">
        - Settings</span></span><br><span data-ttu-id="715a0-287">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-287">
        - TableBindings</span></span><br><span data-ttu-id="715a0-288">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-288">
        - TableCoercion</span></span><br><span data-ttu-id="715a0-289">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-289">
        - TextBindings</span></span><br><span data-ttu-id="715a0-290">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-290">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-291">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="715a0-291">Office 2019 for Mac</span></span><br><span data-ttu-id="715a0-292">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="715a0-292">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="715a0-293">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="715a0-293">- TaskPane</span></span><br><span data-ttu-id="715a0-294">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="715a0-294">
        - Content</span></span><br><span data-ttu-id="715a0-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="715a0-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="715a0-296">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-296">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="715a0-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="715a0-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="715a0-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="715a0-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="715a0-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="715a0-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="715a0-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="715a0-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="715a0-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="715a0-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="715a0-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="715a0-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="715a0-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="715a0-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="715a0-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="715a0-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="715a0-306">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-306">- BindingEvents</span></span><br><span data-ttu-id="715a0-307">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="715a0-307">
        - CompressedFile</span></span><br><span data-ttu-id="715a0-308">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-308">
        - DocumentEvents</span></span><br><span data-ttu-id="715a0-309">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="715a0-309">
        - File</span></span><br><span data-ttu-id="715a0-310">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-310">
        - MatrixBindings</span></span><br><span data-ttu-id="715a0-311">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-311">
        - MatrixCoercion</span></span><br><span data-ttu-id="715a0-312">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="715a0-312">
        - PdfFile</span></span><br><span data-ttu-id="715a0-313">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="715a0-313">
        - Selection</span></span><br><span data-ttu-id="715a0-314">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="715a0-314">
        - Settings</span></span><br><span data-ttu-id="715a0-315">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-315">
        - TableBindings</span></span><br><span data-ttu-id="715a0-316">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-316">
        - TableCoercion</span></span><br><span data-ttu-id="715a0-317">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-317">
        - TextBindings</span></span><br><span data-ttu-id="715a0-318">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-318">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-319">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="715a0-319">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="715a0-320">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="715a0-320">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="715a0-321">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="715a0-321">- TaskPane</span></span><br><span data-ttu-id="715a0-322">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="715a0-322">
        - Content</span></span></td>
    <td><span data-ttu-id="715a0-323">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-323">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="715a0-324">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="715a0-324">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="715a0-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="715a0-326">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-326">- BindingEvents</span></span><br><span data-ttu-id="715a0-327">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="715a0-327">
        - CompressedFile</span></span><br><span data-ttu-id="715a0-328">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-328">
        - DocumentEvents</span></span><br><span data-ttu-id="715a0-329">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="715a0-329">
        - File</span></span><br><span data-ttu-id="715a0-330">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-330">
        - MatrixBindings</span></span><br><span data-ttu-id="715a0-331">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-331">
        - MatrixCoercion</span></span><br><span data-ttu-id="715a0-332">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="715a0-332">
        - PdfFile</span></span><br><span data-ttu-id="715a0-333">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="715a0-333">
        - Selection</span></span><br><span data-ttu-id="715a0-334">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="715a0-334">
        - Settings</span></span><br><span data-ttu-id="715a0-335">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-335">
        - TableBindings</span></span><br><span data-ttu-id="715a0-336">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-336">
        - TableCoercion</span></span><br><span data-ttu-id="715a0-337">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-337">
        - TextBindings</span></span><br><span data-ttu-id="715a0-338">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-338">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="715a0-339">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="715a0-339">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="715a0-340">Funções Personalizadas</span><span class="sxs-lookup"><span data-stu-id="715a0-340">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="715a0-341">Plataforma</span><span class="sxs-lookup"><span data-stu-id="715a0-341">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="715a0-342">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="715a0-342">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="715a0-343">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="715a0-343">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="715a0-344"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="715a0-344"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-345">Office na Web</span><span class="sxs-lookup"><span data-stu-id="715a0-345">Office on the web</span></span></td>
    <td><span data-ttu-id="715a0-346">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="715a0-346">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="715a0-347">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-347">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-348">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="715a0-348">Office on Windows</span></span><br><span data-ttu-id="715a0-349">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="715a0-349">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="715a0-350">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="715a0-350">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="715a0-351">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-351">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-352">Office para Mac</span><span class="sxs-lookup"><span data-stu-id="715a0-352">Office for Mac</span></span><br><span data-ttu-id="715a0-353">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="715a0-353">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="715a0-354">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="715a0-354">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="715a0-355">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-355">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="715a0-356">Outlook</span><span class="sxs-lookup"><span data-stu-id="715a0-356">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="715a0-357">Plataforma</span><span class="sxs-lookup"><span data-stu-id="715a0-357">Platform</span></span></th>
    <th><span data-ttu-id="715a0-358">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="715a0-358">Extension points</span></span></th>
    <th><span data-ttu-id="715a0-359">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="715a0-359">API requirement sets</span></span></th>
    <th><span data-ttu-id="715a0-360"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="715a0-360"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-361">Office na Web</span><span class="sxs-lookup"><span data-stu-id="715a0-361">Office on the web</span></span><br><span data-ttu-id="715a0-362">(moderno)</span><span class="sxs-lookup"><span data-stu-id="715a0-362">Modern</span></span></td>
    <td> <span data-ttu-id="715a0-363">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="715a0-363">- Mail Read</span></span><br><span data-ttu-id="715a0-364">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="715a0-364">
      - Mail Compose</span></span><br><span data-ttu-id="715a0-365">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="715a0-365">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="715a0-366">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-366">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="715a0-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="715a0-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="715a0-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="715a0-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="715a0-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="715a0-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="715a0-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="715a0-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="715a0-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="715a0-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="715a0-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="715a0-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="715a0-373">Não disponível</span><span class="sxs-lookup"><span data-stu-id="715a0-373">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-374">Office na Web</span><span class="sxs-lookup"><span data-stu-id="715a0-374">Office on the web</span></span><br><span data-ttu-id="715a0-375">(clássico)</span><span class="sxs-lookup"><span data-stu-id="715a0-375">Classic.</span></span></td>
    <td> <span data-ttu-id="715a0-376">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="715a0-376">- Mail Read</span></span><br><span data-ttu-id="715a0-377">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="715a0-377">
      - Mail Compose</span></span><br><span data-ttu-id="715a0-378">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="715a0-378">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="715a0-379">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-379">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="715a0-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="715a0-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="715a0-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="715a0-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="715a0-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="715a0-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="715a0-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="715a0-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="715a0-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="715a0-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="715a0-385">Não disponível</span><span class="sxs-lookup"><span data-stu-id="715a0-385">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-386">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="715a0-386">Office on Windows</span></span><br><span data-ttu-id="715a0-387">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="715a0-387">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="715a0-388">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="715a0-388">- Mail Read</span></span><br><span data-ttu-id="715a0-389">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="715a0-389">
      - Mail Compose</span></span><br><span data-ttu-id="715a0-390">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="715a0-390">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="715a0-391">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="715a0-391">
      - Modules</span></span></td>
    <td> <span data-ttu-id="715a0-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="715a0-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="715a0-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="715a0-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="715a0-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="715a0-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="715a0-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="715a0-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="715a0-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="715a0-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="715a0-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="715a0-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="715a0-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="715a0-399">Não disponível</span><span class="sxs-lookup"><span data-stu-id="715a0-399">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-400">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="715a0-400">Office 2019 on Windows</span></span><br><span data-ttu-id="715a0-401">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="715a0-401">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="715a0-402">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="715a0-402">- Mail Read</span></span><br><span data-ttu-id="715a0-403">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="715a0-403">
      - Mail Compose</span></span><br><span data-ttu-id="715a0-404">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="715a0-404">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="715a0-405">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="715a0-405">
      - Modules</span></span></td>
    <td> <span data-ttu-id="715a0-406">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-406">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="715a0-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="715a0-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="715a0-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="715a0-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="715a0-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="715a0-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="715a0-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="715a0-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="715a0-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="715a0-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="715a0-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="715a0-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="715a0-413">Não disponível</span><span class="sxs-lookup"><span data-stu-id="715a0-413">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-414">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="715a0-414">Office 2016 on Windows</span></span><br><span data-ttu-id="715a0-415">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="715a0-415">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="715a0-416">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="715a0-416">- Mail Read</span></span><br><span data-ttu-id="715a0-417">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="715a0-417">
      - Mail Compose</span></span><br><span data-ttu-id="715a0-418">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="715a0-418">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="715a0-419">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="715a0-419">
      - Modules</span></span></td>
    <td> <span data-ttu-id="715a0-420">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-420">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="715a0-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="715a0-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="715a0-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="715a0-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="715a0-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="715a0-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="715a0-424">Não disponível</span><span class="sxs-lookup"><span data-stu-id="715a0-424">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-425">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="715a0-425">Office 2013 on Windows</span></span><br><span data-ttu-id="715a0-426">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="715a0-426">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="715a0-427">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="715a0-427">- Mail Read</span></span><br><span data-ttu-id="715a0-428">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="715a0-428">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="715a0-429">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-429">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="715a0-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="715a0-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="715a0-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="715a0-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="715a0-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="715a0-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="715a0-433">Não disponível</span><span class="sxs-lookup"><span data-stu-id="715a0-433">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-434">Office no iOS</span><span class="sxs-lookup"><span data-stu-id="715a0-434">Office apps on iOS</span></span><br><span data-ttu-id="715a0-435">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="715a0-435">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="715a0-436">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="715a0-436">- Mail Read</span></span><br><span data-ttu-id="715a0-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="715a0-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="715a0-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="715a0-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="715a0-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="715a0-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="715a0-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="715a0-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="715a0-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="715a0-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="715a0-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="715a0-443">Não disponível</span><span class="sxs-lookup"><span data-stu-id="715a0-443">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-444">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="715a0-444">Office apps on Mac</span></span><br><span data-ttu-id="715a0-445">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="715a0-445">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="715a0-446">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="715a0-446">- Mail Read</span></span><br><span data-ttu-id="715a0-447">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="715a0-447">
      - Mail Compose</span></span><br><span data-ttu-id="715a0-448">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="715a0-448">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="715a0-449">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-449">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="715a0-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="715a0-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="715a0-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="715a0-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="715a0-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="715a0-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="715a0-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="715a0-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="715a0-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="715a0-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="715a0-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="715a0-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="715a0-456">Não disponível</span><span class="sxs-lookup"><span data-stu-id="715a0-456">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-457">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="715a0-457">Office 2019 for Mac</span></span><br><span data-ttu-id="715a0-458">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="715a0-458">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="715a0-459">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="715a0-459">- Mail Read</span></span><br><span data-ttu-id="715a0-460">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="715a0-460">
      - Mail Compose</span></span><br><span data-ttu-id="715a0-461">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="715a0-461">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="715a0-462">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-462">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="715a0-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="715a0-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="715a0-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="715a0-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="715a0-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="715a0-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="715a0-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="715a0-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="715a0-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="715a0-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="715a0-468">Não disponível</span><span class="sxs-lookup"><span data-stu-id="715a0-468">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-469">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="715a0-469">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="715a0-470">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="715a0-470">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="715a0-471">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="715a0-471">- Mail Read</span></span><br><span data-ttu-id="715a0-472">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="715a0-472">
      - Mail Compose</span></span><br><span data-ttu-id="715a0-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="715a0-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="715a0-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="715a0-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="715a0-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="715a0-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="715a0-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="715a0-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="715a0-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="715a0-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="715a0-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="715a0-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="715a0-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="715a0-480">Não disponível</span><span class="sxs-lookup"><span data-stu-id="715a0-480">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-481">Outlook no Android</span><span class="sxs-lookup"><span data-stu-id="715a0-481">Office apps on Android</span></span><br><span data-ttu-id="715a0-482">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="715a0-482">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="715a0-483">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="715a0-483">- Mail Read</span></span><br><span data-ttu-id="715a0-484">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="715a0-484">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="715a0-485">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-485">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="715a0-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="715a0-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="715a0-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="715a0-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="715a0-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="715a0-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="715a0-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="715a0-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="715a0-490">Não disponível</span><span class="sxs-lookup"><span data-stu-id="715a0-490">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="715a0-491">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="715a0-491">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="715a0-492">Word</span><span class="sxs-lookup"><span data-stu-id="715a0-492">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="715a0-493">Plataforma</span><span class="sxs-lookup"><span data-stu-id="715a0-493">Platform</span></span></th>
    <th><span data-ttu-id="715a0-494">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="715a0-494">Extension points</span></span></th>
    <th><span data-ttu-id="715a0-495">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="715a0-495">API requirement sets</span></span></th>
    <th><span data-ttu-id="715a0-496"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="715a0-496"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-497">Office na Web</span><span class="sxs-lookup"><span data-stu-id="715a0-497">Office on the web</span></span></td>
    <td> <span data-ttu-id="715a0-498">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="715a0-498">- TaskPane</span></span><br><span data-ttu-id="715a0-499">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="715a0-499">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="715a0-500">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-500">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="715a0-501">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="715a0-501">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="715a0-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="715a0-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="715a0-503">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-503">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="715a0-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="715a0-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="715a0-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="715a0-506">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-506">- BindingEvents</span></span><br><span data-ttu-id="715a0-507">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="715a0-507">
         - CustomXmlParts</span></span><br><span data-ttu-id="715a0-508">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-508">
         - DocumentEvents</span></span><br><span data-ttu-id="715a0-509">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="715a0-509">
         - File</span></span><br><span data-ttu-id="715a0-510">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-510">
         - HtmlCoercion</span></span><br><span data-ttu-id="715a0-511">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-511">
         - MatrixBindings</span></span><br><span data-ttu-id="715a0-512">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-512">
         - MatrixCoercion</span></span><br><span data-ttu-id="715a0-513">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-513">
         - OoxmlCoercion</span></span><br><span data-ttu-id="715a0-514">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="715a0-514">
         - PdfFile</span></span><br><span data-ttu-id="715a0-515">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="715a0-515">
         - Selection</span></span><br><span data-ttu-id="715a0-516">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="715a0-516">
         - Settings</span></span><br><span data-ttu-id="715a0-517">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-517">
         - TableBindings</span></span><br><span data-ttu-id="715a0-518">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-518">
         - TableCoercion</span></span><br><span data-ttu-id="715a0-519">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-519">
         - TextBindings</span></span><br><span data-ttu-id="715a0-520">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-520">
         - TextCoercion</span></span><br><span data-ttu-id="715a0-521">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="715a0-521">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-522">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="715a0-522">Office on Windows</span></span><br><span data-ttu-id="715a0-523">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="715a0-523">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="715a0-524">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="715a0-524">- TaskPane</span></span><br><span data-ttu-id="715a0-525">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="715a0-525">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="715a0-526">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-526">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="715a0-527">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="715a0-527">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="715a0-528">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="715a0-528">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="715a0-529">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-529">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="715a0-530">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-530">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="715a0-531">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="715a0-531">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="715a0-532">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-532">- BindingEvents</span></span><br><span data-ttu-id="715a0-533">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="715a0-533">
         - CompressedFile</span></span><br><span data-ttu-id="715a0-534">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="715a0-534">
         - CustomXmlParts</span></span><br><span data-ttu-id="715a0-535">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-535">
         - DocumentEvents</span></span><br><span data-ttu-id="715a0-536">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="715a0-536">
         - File</span></span><br><span data-ttu-id="715a0-537">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-537">
         - HtmlCoercion</span></span><br><span data-ttu-id="715a0-538">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-538">
         - MatrixBindings</span></span><br><span data-ttu-id="715a0-539">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-539">
         - MatrixCoercion</span></span><br><span data-ttu-id="715a0-540">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-540">
         - OoxmlCoercion</span></span><br><span data-ttu-id="715a0-541">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="715a0-541">
         - PdfFile</span></span><br><span data-ttu-id="715a0-542">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="715a0-542">
         - Selection</span></span><br><span data-ttu-id="715a0-543">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="715a0-543">
         - Settings</span></span><br><span data-ttu-id="715a0-544">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-544">
         - TableBindings</span></span><br><span data-ttu-id="715a0-545">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-545">
         - TableCoercion</span></span><br><span data-ttu-id="715a0-546">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-546">
         - TextBindings</span></span><br><span data-ttu-id="715a0-547">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-547">
         - TextCoercion</span></span><br><span data-ttu-id="715a0-548">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="715a0-548">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-549">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="715a0-549">Office 2019 on Windows</span></span><br><span data-ttu-id="715a0-550">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="715a0-550">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="715a0-551">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="715a0-551">- TaskPane</span></span><br><span data-ttu-id="715a0-552">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="715a0-552">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="715a0-553">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-553">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="715a0-554">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="715a0-554">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="715a0-555">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="715a0-555">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="715a0-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="715a0-557">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-557">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="715a0-558">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-558">- BindingEvents</span></span><br><span data-ttu-id="715a0-559">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="715a0-559">
         - CompressedFile</span></span><br><span data-ttu-id="715a0-560">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="715a0-560">
         - CustomXmlParts</span></span><br><span data-ttu-id="715a0-561">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-561">
         - DocumentEvents</span></span><br><span data-ttu-id="715a0-562">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="715a0-562">
         - File</span></span><br><span data-ttu-id="715a0-563">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-563">
         - HtmlCoercion</span></span><br><span data-ttu-id="715a0-564">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-564">
         - MatrixBindings</span></span><br><span data-ttu-id="715a0-565">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-565">
         - MatrixCoercion</span></span><br><span data-ttu-id="715a0-566">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-566">
         - OoxmlCoercion</span></span><br><span data-ttu-id="715a0-567">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="715a0-567">
         - PdfFile</span></span><br><span data-ttu-id="715a0-568">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="715a0-568">
         - Selection</span></span><br><span data-ttu-id="715a0-569">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="715a0-569">
         - Settings</span></span><br><span data-ttu-id="715a0-570">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-570">
         - TableBindings</span></span><br><span data-ttu-id="715a0-571">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-571">
         - TableCoercion</span></span><br><span data-ttu-id="715a0-572">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-572">
         - TextBindings</span></span><br><span data-ttu-id="715a0-573">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-573">
         - TextCoercion</span></span><br><span data-ttu-id="715a0-574">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="715a0-574">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-575">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="715a0-575">Office 2016 on Windows</span></span><br><span data-ttu-id="715a0-576">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="715a0-576">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="715a0-577">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="715a0-577">- TaskPane</span></span></td>
    <td> <span data-ttu-id="715a0-578">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-578">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="715a0-579">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="715a0-579">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="715a0-580">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-580">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="715a0-581">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-581">- BindingEvents</span></span><br><span data-ttu-id="715a0-582">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="715a0-582">
         - CompressedFile</span></span><br><span data-ttu-id="715a0-583">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="715a0-583">
         - CustomXmlParts</span></span><br><span data-ttu-id="715a0-584">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-584">
         - DocumentEvents</span></span><br><span data-ttu-id="715a0-585">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="715a0-585">
         - File</span></span><br><span data-ttu-id="715a0-586">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-586">
         - HtmlCoercion</span></span><br><span data-ttu-id="715a0-587">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-587">
         - MatrixBindings</span></span><br><span data-ttu-id="715a0-588">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-588">
         - MatrixCoercion</span></span><br><span data-ttu-id="715a0-589">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-589">
         - OoxmlCoercion</span></span><br><span data-ttu-id="715a0-590">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="715a0-590">
         - PdfFile</span></span><br><span data-ttu-id="715a0-591">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="715a0-591">
         - Selection</span></span><br><span data-ttu-id="715a0-592">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="715a0-592">
         - Settings</span></span><br><span data-ttu-id="715a0-593">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-593">
         - TableBindings</span></span><br><span data-ttu-id="715a0-594">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-594">
         - TableCoercion</span></span><br><span data-ttu-id="715a0-595">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-595">
         - TextBindings</span></span><br><span data-ttu-id="715a0-596">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-596">
         - TextCoercion</span></span><br><span data-ttu-id="715a0-597">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="715a0-597">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-598">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="715a0-598">Office 2013 on Windows</span></span><br><span data-ttu-id="715a0-599">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="715a0-599">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="715a0-600">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="715a0-600">- TaskPane</span></span></td>
    <td> <span data-ttu-id="715a0-601">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="715a0-601">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="715a0-602">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-602">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="715a0-603">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-603">- BindingEvents</span></span><br><span data-ttu-id="715a0-604">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="715a0-604">
         - CompressedFile</span></span><br><span data-ttu-id="715a0-605">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="715a0-605">
         - CustomXmlParts</span></span><br><span data-ttu-id="715a0-606">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-606">
         - DocumentEvents</span></span><br><span data-ttu-id="715a0-607">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="715a0-607">
         - File</span></span><br><span data-ttu-id="715a0-608">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-608">
         - HtmlCoercion</span></span><br><span data-ttu-id="715a0-609">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-609">
         - MatrixBindings</span></span><br><span data-ttu-id="715a0-610">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-610">
         - MatrixCoercion</span></span><br><span data-ttu-id="715a0-611">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-611">
         - OoxmlCoercion</span></span><br><span data-ttu-id="715a0-612">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="715a0-612">
         - PdfFile</span></span><br><span data-ttu-id="715a0-613">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="715a0-613">
         - Selection</span></span><br><span data-ttu-id="715a0-614">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="715a0-614">
         - Settings</span></span><br><span data-ttu-id="715a0-615">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-615">
         - TableBindings</span></span><br><span data-ttu-id="715a0-616">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-616">
         - TableCoercion</span></span><br><span data-ttu-id="715a0-617">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-617">
         - TextBindings</span></span><br><span data-ttu-id="715a0-618">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-618">
         - TextCoercion</span></span><br><span data-ttu-id="715a0-619">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="715a0-619">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-620">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="715a0-620">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="715a0-621">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="715a0-621">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="715a0-622">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="715a0-622">- TaskPane</span></span></td>
    <td> <span data-ttu-id="715a0-623">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-623">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="715a0-624">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="715a0-624">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="715a0-625">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="715a0-625">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="715a0-626">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-626">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="715a0-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="715a0-628">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-628">- BindingEvents</span></span><br><span data-ttu-id="715a0-629">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="715a0-629">
         - CompressedFile</span></span><br><span data-ttu-id="715a0-630">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="715a0-630">
         - CustomXmlParts</span></span><br><span data-ttu-id="715a0-631">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-631">
         - DocumentEvents</span></span><br><span data-ttu-id="715a0-632">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="715a0-632">
         - File</span></span><br><span data-ttu-id="715a0-633">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-633">
         - HtmlCoercion</span></span><br><span data-ttu-id="715a0-634">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-634">
         - MatrixBindings</span></span><br><span data-ttu-id="715a0-635">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-635">
         - MatrixCoercion</span></span><br><span data-ttu-id="715a0-636">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-636">
         - OoxmlCoercion</span></span><br><span data-ttu-id="715a0-637">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="715a0-637">
         - PdfFile</span></span><br><span data-ttu-id="715a0-638">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="715a0-638">
         - Selection</span></span><br><span data-ttu-id="715a0-639">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="715a0-639">
         - Settings</span></span><br><span data-ttu-id="715a0-640">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-640">
         - TableBindings</span></span><br><span data-ttu-id="715a0-641">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-641">
         - TableCoercion</span></span><br><span data-ttu-id="715a0-642">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-642">
         - TextBindings</span></span><br><span data-ttu-id="715a0-643">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-643">
         - TextCoercion</span></span><br><span data-ttu-id="715a0-644">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="715a0-644">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-645">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="715a0-645">Office apps on Mac</span></span><br><span data-ttu-id="715a0-646">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="715a0-646">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="715a0-647">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="715a0-647">- TaskPane</span></span><br><span data-ttu-id="715a0-648">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="715a0-648">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="715a0-649">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-649">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="715a0-650">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="715a0-650">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="715a0-651">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="715a0-651">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="715a0-652">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-652">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="715a0-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="715a0-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="715a0-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="715a0-655">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-655">- BindingEvents</span></span><br><span data-ttu-id="715a0-656">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="715a0-656">
         - CompressedFile</span></span><br><span data-ttu-id="715a0-657">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="715a0-657">
         - CustomXmlParts</span></span><br><span data-ttu-id="715a0-658">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-658">
         - DocumentEvents</span></span><br><span data-ttu-id="715a0-659">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="715a0-659">
         - File</span></span><br><span data-ttu-id="715a0-660">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-660">
         - HtmlCoercion</span></span><br><span data-ttu-id="715a0-661">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-661">
         - MatrixBindings</span></span><br><span data-ttu-id="715a0-662">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-662">
         - MatrixCoercion</span></span><br><span data-ttu-id="715a0-663">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-663">
         - OoxmlCoercion</span></span><br><span data-ttu-id="715a0-664">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="715a0-664">
         - PdfFile</span></span><br><span data-ttu-id="715a0-665">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="715a0-665">
         - Selection</span></span><br><span data-ttu-id="715a0-666">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="715a0-666">
         - Settings</span></span><br><span data-ttu-id="715a0-667">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-667">
         - TableBindings</span></span><br><span data-ttu-id="715a0-668">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-668">
         - TableCoercion</span></span><br><span data-ttu-id="715a0-669">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-669">
         - TextBindings</span></span><br><span data-ttu-id="715a0-670">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-670">
         - TextCoercion</span></span><br><span data-ttu-id="715a0-671">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="715a0-671">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-672">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="715a0-672">Office 2019 for Mac</span></span><br><span data-ttu-id="715a0-673">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="715a0-673">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="715a0-674">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="715a0-674">- TaskPane</span></span><br><span data-ttu-id="715a0-675">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="715a0-675">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="715a0-676">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-676">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="715a0-677">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="715a0-677">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="715a0-678">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="715a0-678">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="715a0-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="715a0-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="715a0-681">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-681">- BindingEvents</span></span><br><span data-ttu-id="715a0-682">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="715a0-682">
         - CompressedFile</span></span><br><span data-ttu-id="715a0-683">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="715a0-683">
         - CustomXmlParts</span></span><br><span data-ttu-id="715a0-684">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-684">
         - DocumentEvents</span></span><br><span data-ttu-id="715a0-685">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="715a0-685">
         - File</span></span><br><span data-ttu-id="715a0-686">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-686">
         - HtmlCoercion</span></span><br><span data-ttu-id="715a0-687">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-687">
         - MatrixBindings</span></span><br><span data-ttu-id="715a0-688">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-688">
         - MatrixCoercion</span></span><br><span data-ttu-id="715a0-689">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-689">
         - OoxmlCoercion</span></span><br><span data-ttu-id="715a0-690">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="715a0-690">
         - PdfFile</span></span><br><span data-ttu-id="715a0-691">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="715a0-691">
         - Selection</span></span><br><span data-ttu-id="715a0-692">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="715a0-692">
         - Settings</span></span><br><span data-ttu-id="715a0-693">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-693">
         - TableBindings</span></span><br><span data-ttu-id="715a0-694">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-694">
         - TableCoercion</span></span><br><span data-ttu-id="715a0-695">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-695">
         - TextBindings</span></span><br><span data-ttu-id="715a0-696">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-696">
         - TextCoercion</span></span><br><span data-ttu-id="715a0-697">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="715a0-697">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-698">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="715a0-698">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="715a0-699">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="715a0-699">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="715a0-700">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="715a0-700">- TaskPane</span></span></td>
    <td> <span data-ttu-id="715a0-701">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-701">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="715a0-702">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="715a0-702">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="715a0-703">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-703">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="715a0-704">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-704">- BindingEvents</span></span><br><span data-ttu-id="715a0-705">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="715a0-705">
         - CompressedFile</span></span><br><span data-ttu-id="715a0-706">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="715a0-706">
         - CustomXmlParts</span></span><br><span data-ttu-id="715a0-707">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-707">
         - DocumentEvents</span></span><br><span data-ttu-id="715a0-708">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="715a0-708">
         - File</span></span><br><span data-ttu-id="715a0-709">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-709">
         - HtmlCoercion</span></span><br><span data-ttu-id="715a0-710">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-710">
         - MatrixBindings</span></span><br><span data-ttu-id="715a0-711">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-711">
         - MatrixCoercion</span></span><br><span data-ttu-id="715a0-712">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-712">
         - OoxmlCoercion</span></span><br><span data-ttu-id="715a0-713">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="715a0-713">
         - PdfFile</span></span><br><span data-ttu-id="715a0-714">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="715a0-714">
         - Selection</span></span><br><span data-ttu-id="715a0-715">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="715a0-715">
         - Settings</span></span><br><span data-ttu-id="715a0-716">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-716">
         - TableBindings</span></span><br><span data-ttu-id="715a0-717">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-717">
         - TableCoercion</span></span><br><span data-ttu-id="715a0-718">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="715a0-718">
         - TextBindings</span></span><br><span data-ttu-id="715a0-719">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-719">
         - TextCoercion</span></span><br><span data-ttu-id="715a0-720">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="715a0-720">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="715a0-721">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="715a0-721">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="715a0-722">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="715a0-722">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="715a0-723">Plataforma</span><span class="sxs-lookup"><span data-stu-id="715a0-723">Platform</span></span></th>
    <th><span data-ttu-id="715a0-724">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="715a0-724">Extension points</span></span></th>
    <th><span data-ttu-id="715a0-725">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="715a0-725">API requirement sets</span></span></th>
    <th><span data-ttu-id="715a0-726"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="715a0-726"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-727">Office na Web</span><span class="sxs-lookup"><span data-stu-id="715a0-727">Office on the web</span></span></td>
    <td> <span data-ttu-id="715a0-728">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="715a0-728">- Content</span></span><br><span data-ttu-id="715a0-729">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="715a0-729">
         - TaskPane</span></span><br><span data-ttu-id="715a0-730">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="715a0-730">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="715a0-731">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-731">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="715a0-732">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-732">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="715a0-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="715a0-734">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="715a0-734">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="715a0-735">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="715a0-735">- ActiveView</span></span><br><span data-ttu-id="715a0-736">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="715a0-736">
         - CompressedFile</span></span><br><span data-ttu-id="715a0-737">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-737">
         - DocumentEvents</span></span><br><span data-ttu-id="715a0-738">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="715a0-738">
         - File</span></span><br><span data-ttu-id="715a0-739">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="715a0-739">
         - PdfFile</span></span><br><span data-ttu-id="715a0-740">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="715a0-740">
         - Selection</span></span><br><span data-ttu-id="715a0-741">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="715a0-741">
         - Settings</span></span><br><span data-ttu-id="715a0-742">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-742">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-743">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="715a0-743">Office on Windows</span></span><br><span data-ttu-id="715a0-744">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="715a0-744">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="715a0-745">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="715a0-745">- Content</span></span><br><span data-ttu-id="715a0-746">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="715a0-746">
         - TaskPane</span></span><br><span data-ttu-id="715a0-747">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="715a0-747">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="715a0-748">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-748">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="715a0-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="715a0-750">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-750">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="715a0-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="715a0-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="715a0-752">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="715a0-752">- ActiveView</span></span><br><span data-ttu-id="715a0-753">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="715a0-753">
         - CompressedFile</span></span><br><span data-ttu-id="715a0-754">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-754">
         - DocumentEvents</span></span><br><span data-ttu-id="715a0-755">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="715a0-755">
         - File</span></span><br><span data-ttu-id="715a0-756">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="715a0-756">
         - PdfFile</span></span><br><span data-ttu-id="715a0-757">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="715a0-757">
         - Selection</span></span><br><span data-ttu-id="715a0-758">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="715a0-758">
         - Settings</span></span><br><span data-ttu-id="715a0-759">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-759">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-760">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="715a0-760">Office 2019 on Windows</span></span><br><span data-ttu-id="715a0-761">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="715a0-761">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="715a0-762">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="715a0-762">- Content</span></span><br><span data-ttu-id="715a0-763">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="715a0-763">
         - TaskPane</span></span><br><span data-ttu-id="715a0-764">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="715a0-764">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="715a0-765">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-765">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="715a0-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="715a0-767">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="715a0-767">- ActiveView</span></span><br><span data-ttu-id="715a0-768">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="715a0-768">
         - CompressedFile</span></span><br><span data-ttu-id="715a0-769">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-769">
         - DocumentEvents</span></span><br><span data-ttu-id="715a0-770">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="715a0-770">
         - File</span></span><br><span data-ttu-id="715a0-771">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="715a0-771">
         - PdfFile</span></span><br><span data-ttu-id="715a0-772">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="715a0-772">
         - Selection</span></span><br><span data-ttu-id="715a0-773">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="715a0-773">
         - Settings</span></span><br><span data-ttu-id="715a0-774">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-774">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-775">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="715a0-775">Office 2016 on Windows</span></span><br><span data-ttu-id="715a0-776">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="715a0-776">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="715a0-777">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="715a0-777">- Content</span></span><br><span data-ttu-id="715a0-778">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="715a0-778">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="715a0-779">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="715a0-779">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="715a0-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="715a0-781">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="715a0-781">- ActiveView</span></span><br><span data-ttu-id="715a0-782">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="715a0-782">
         - CompressedFile</span></span><br><span data-ttu-id="715a0-783">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-783">
         - DocumentEvents</span></span><br><span data-ttu-id="715a0-784">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="715a0-784">
         - File</span></span><br><span data-ttu-id="715a0-785">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="715a0-785">
         - PdfFile</span></span><br><span data-ttu-id="715a0-786">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="715a0-786">
         - Selection</span></span><br><span data-ttu-id="715a0-787">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="715a0-787">
         - Settings</span></span><br><span data-ttu-id="715a0-788">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-788">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-789">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="715a0-789">Office 2013 on Windows</span></span><br><span data-ttu-id="715a0-790">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="715a0-790">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="715a0-791">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="715a0-791">- Content</span></span><br><span data-ttu-id="715a0-792">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="715a0-792">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="715a0-793">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="715a0-793">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="715a0-794">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-794">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="715a0-795">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="715a0-795">- ActiveView</span></span><br><span data-ttu-id="715a0-796">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="715a0-796">
         - CompressedFile</span></span><br><span data-ttu-id="715a0-797">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-797">
         - DocumentEvents</span></span><br><span data-ttu-id="715a0-798">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="715a0-798">
         - File</span></span><br><span data-ttu-id="715a0-799">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="715a0-799">
         - PdfFile</span></span><br><span data-ttu-id="715a0-800">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="715a0-800">
         - Selection</span></span><br><span data-ttu-id="715a0-801">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="715a0-801">
         - Settings</span></span><br><span data-ttu-id="715a0-802">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-802">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-803">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="715a0-803">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="715a0-804">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="715a0-804">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="715a0-805">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="715a0-805">- Content</span></span><br><span data-ttu-id="715a0-806">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="715a0-806">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="715a0-807">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-807">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="715a0-808">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-808">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="715a0-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="715a0-810">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="715a0-810">- ActiveView</span></span><br><span data-ttu-id="715a0-811">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="715a0-811">
         - CompressedFile</span></span><br><span data-ttu-id="715a0-812">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-812">
         - DocumentEvents</span></span><br><span data-ttu-id="715a0-813">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="715a0-813">
         - File</span></span><br><span data-ttu-id="715a0-814">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="715a0-814">
         - PdfFile</span></span><br><span data-ttu-id="715a0-815">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="715a0-815">
         - Selection</span></span><br><span data-ttu-id="715a0-816">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="715a0-816">
         - Settings</span></span><br><span data-ttu-id="715a0-817">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-817">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-818">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="715a0-818">Office apps on Mac</span></span><br><span data-ttu-id="715a0-819">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="715a0-819">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="715a0-820">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="715a0-820">- Content</span></span><br><span data-ttu-id="715a0-821">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="715a0-821">
         - TaskPane</span></span><br><span data-ttu-id="715a0-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="715a0-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="715a0-823">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-823">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="715a0-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="715a0-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="715a0-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="715a0-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="715a0-827">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="715a0-827">- ActiveView</span></span><br><span data-ttu-id="715a0-828">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="715a0-828">
         - CompressedFile</span></span><br><span data-ttu-id="715a0-829">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-829">
         - DocumentEvents</span></span><br><span data-ttu-id="715a0-830">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="715a0-830">
         - File</span></span><br><span data-ttu-id="715a0-831">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="715a0-831">
         - PdfFile</span></span><br><span data-ttu-id="715a0-832">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="715a0-832">
         - Selection</span></span><br><span data-ttu-id="715a0-833">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="715a0-833">
         - Settings</span></span><br><span data-ttu-id="715a0-834">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-834">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-835">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="715a0-835">Office 2019 for Mac</span></span><br><span data-ttu-id="715a0-836">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="715a0-836">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="715a0-837">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="715a0-837">- Content</span></span><br><span data-ttu-id="715a0-838">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="715a0-838">
         - TaskPane</span></span><br><span data-ttu-id="715a0-839">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="715a0-839">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="715a0-840">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-840">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="715a0-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="715a0-842">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="715a0-842">- ActiveView</span></span><br><span data-ttu-id="715a0-843">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="715a0-843">
         - CompressedFile</span></span><br><span data-ttu-id="715a0-844">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-844">
         - DocumentEvents</span></span><br><span data-ttu-id="715a0-845">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="715a0-845">
         - File</span></span><br><span data-ttu-id="715a0-846">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="715a0-846">
         - PdfFile</span></span><br><span data-ttu-id="715a0-847">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="715a0-847">
         - Selection</span></span><br><span data-ttu-id="715a0-848">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="715a0-848">
         - Settings</span></span><br><span data-ttu-id="715a0-849">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-849">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-850">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="715a0-850">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="715a0-851">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="715a0-851">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="715a0-852">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="715a0-852">- Content</span></span><br><span data-ttu-id="715a0-853">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="715a0-853">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="715a0-854">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="715a0-854">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="715a0-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="715a0-856">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="715a0-856">- ActiveView</span></span><br><span data-ttu-id="715a0-857">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="715a0-857">
         - CompressedFile</span></span><br><span data-ttu-id="715a0-858">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-858">
         - DocumentEvents</span></span><br><span data-ttu-id="715a0-859">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="715a0-859">
         - File</span></span><br><span data-ttu-id="715a0-860">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="715a0-860">
         - PdfFile</span></span><br><span data-ttu-id="715a0-861">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="715a0-861">
         - Selection</span></span><br><span data-ttu-id="715a0-862">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="715a0-862">
         - Settings</span></span><br><span data-ttu-id="715a0-863">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-863">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="715a0-864">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="715a0-864">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="715a0-865">OneNote</span><span class="sxs-lookup"><span data-stu-id="715a0-865">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="715a0-866">Plataforma</span><span class="sxs-lookup"><span data-stu-id="715a0-866">Platform</span></span></th>
    <th><span data-ttu-id="715a0-867">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="715a0-867">Extension points</span></span></th>
    <th><span data-ttu-id="715a0-868">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="715a0-868">API requirement sets</span></span></th>
    <th><span data-ttu-id="715a0-869"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="715a0-869"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-870">Office na Web</span><span class="sxs-lookup"><span data-stu-id="715a0-870">Office on the web</span></span></td>
    <td> <span data-ttu-id="715a0-871">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="715a0-871">- Content</span></span><br><span data-ttu-id="715a0-872">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="715a0-872">
         - TaskPane</span></span><br><span data-ttu-id="715a0-873">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="715a0-873">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="715a0-874">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-874">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="715a0-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="715a0-876">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-876">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="715a0-877">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="715a0-877">- DocumentEvents</span></span><br><span data-ttu-id="715a0-878">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-878">
         - HtmlCoercion</span></span><br><span data-ttu-id="715a0-879">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="715a0-879">
         - Settings</span></span><br><span data-ttu-id="715a0-880">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-880">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="715a0-881">Project</span><span class="sxs-lookup"><span data-stu-id="715a0-881">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="715a0-882">Plataforma</span><span class="sxs-lookup"><span data-stu-id="715a0-882">Platform</span></span></th>
    <th><span data-ttu-id="715a0-883">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="715a0-883">Extension points</span></span></th>
    <th><span data-ttu-id="715a0-884">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="715a0-884">API requirement sets</span></span></th>
    <th><span data-ttu-id="715a0-885"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="715a0-885"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-886">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="715a0-886">Office 2019 on Windows</span></span><br><span data-ttu-id="715a0-887">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="715a0-887">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="715a0-888">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="715a0-888">- TaskPane</span></span></td>
    <td> <span data-ttu-id="715a0-889">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-889">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="715a0-890">- Seleção</span><span class="sxs-lookup"><span data-stu-id="715a0-890">- Selection</span></span><br><span data-ttu-id="715a0-891">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-891">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-892">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="715a0-892">Office 2016 on Windows</span></span><br><span data-ttu-id="715a0-893">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="715a0-893">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="715a0-894">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="715a0-894">- TaskPane</span></span></td>
    <td> <span data-ttu-id="715a0-895">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-895">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="715a0-896">- Seleção</span><span class="sxs-lookup"><span data-stu-id="715a0-896">- Selection</span></span><br><span data-ttu-id="715a0-897">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-897">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="715a0-898">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="715a0-898">Office 2013 on Windows</span></span><br><span data-ttu-id="715a0-899">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="715a0-899">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="715a0-900">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="715a0-900">- TaskPane</span></span></td>
    <td> <span data-ttu-id="715a0-901">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="715a0-901">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="715a0-902">- Seleção</span><span class="sxs-lookup"><span data-stu-id="715a0-902">- Selection</span></span><br><span data-ttu-id="715a0-903">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="715a0-903">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="715a0-904">Confira também</span><span class="sxs-lookup"><span data-stu-id="715a0-904">See also</span></span>

- [<span data-ttu-id="715a0-905">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="715a0-905">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="715a0-906">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="715a0-906">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="715a0-907">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="715a0-907">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="715a0-908">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="715a0-908">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="715a0-909">Referência da API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="715a0-909">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="715a0-910">Histórico de atualizações do Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="715a0-910">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="715a0-911">Histórico de atualizações do Office 2016 e 2019 (Clique para Executar)</span><span class="sxs-lookup"><span data-stu-id="715a0-911">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="715a0-912">Histórico de atualizações do Office 2013 (clique para executar)</span><span class="sxs-lookup"><span data-stu-id="715a0-912">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="715a0-913">Histórico de atualizações do Office 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="715a0-913">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="715a0-914">Histórico de atualizações do Outlook 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="715a0-914">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="715a0-915">Histórico de atualizações do Office para Mac</span><span class="sxs-lookup"><span data-stu-id="715a0-915">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
