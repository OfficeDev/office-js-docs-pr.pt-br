---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos com suporte para o Excel, OneNote, Outlook, PowerPoint, Project e Word.
ms.date: 01/23/2020
localization_priority: Priority
ms.openlocfilehash: b30fe872fd89bb02afac99a7838d43d1fbee5464
ms.sourcegitcommit: a0262ea40cd23f221e69bcb0223110f011265d13
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42688602"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="c3d2f-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="c3d2f-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="c3d2f-104">Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="c3d2f-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="c3d2f-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que são compatíveis atualmente com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="c3d2f-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="c3d2f-106">A versão inicial do Office 2016 instalada por meio do MSI apenas contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="c3d2f-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="c3d2f-107">Para saber mais sobre o histórico de atualizações de várias versões do Office, confira a seção[Confira também](#see-also).</span><span class="sxs-lookup"><span data-stu-id="c3d2f-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="c3d2f-108">Excel</span><span class="sxs-lookup"><span data-stu-id="c3d2f-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="c3d2f-109">Plataforma</span><span class="sxs-lookup"><span data-stu-id="c3d2f-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="c3d2f-110">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="c3d2f-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="c3d2f-111">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="c3d2f-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="c3d2f-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-113">Office na Web</span><span class="sxs-lookup"><span data-stu-id="c3d2f-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="c3d2f-114">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c3d2f-114">- TaskPane</span></span><br><span data-ttu-id="c3d2f-115">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-115">
        - Content</span></span><br><span data-ttu-id="c3d2f-116">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="c3d2f-116">
        - Custom Functions</span></span><br><span data-ttu-id="c3d2f-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="c3d2f-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="c3d2f-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c3d2f-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c3d2f-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c3d2f-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c3d2f-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c3d2f-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c3d2f-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c3d2f-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="c3d2f-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="c3d2f-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="c3d2f-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="c3d2f-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-130">
        - BindingEvents</span></span><br><span data-ttu-id="c3d2f-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-131">
        - CompressedFile</span></span><br><span data-ttu-id="c3d2f-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-132">
        - DocumentEvents</span></span><br><span data-ttu-id="c3d2f-133">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-133">
        - File</span></span><br><span data-ttu-id="c3d2f-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-134">
        - MatrixBindings</span></span><br><span data-ttu-id="c3d2f-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="c3d2f-136">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="c3d2f-136">
        - Selection</span></span><br><span data-ttu-id="c3d2f-137">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="c3d2f-137">
        - Settings</span></span><br><span data-ttu-id="c3d2f-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-138">
        - TableBindings</span></span><br><span data-ttu-id="c3d2f-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-139">
        - TableCoercion</span></span><br><span data-ttu-id="c3d2f-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-140">
        - TextBindings</span></span><br><span data-ttu-id="c3d2f-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-142">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="c3d2f-142">Office on Windows</span></span><br><span data-ttu-id="c3d2f-143">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c3d2f-144">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c3d2f-144">- TaskPane</span></span><br><span data-ttu-id="c3d2f-145">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-145">
        - Content</span></span><br><span data-ttu-id="c3d2f-146">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="c3d2f-146">
        - Custom Functions</span></span><br><span data-ttu-id="c3d2f-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="c3d2f-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="c3d2f-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c3d2f-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c3d2f-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c3d2f-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c3d2f-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c3d2f-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c3d2f-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c3d2f-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="c3d2f-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="c3d2f-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c3d2f-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="c3d2f-161">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-161">
        - BindingEvents</span></span><br><span data-ttu-id="c3d2f-162">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-162">
        - CompressedFile</span></span><br><span data-ttu-id="c3d2f-163">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-163">
        - DocumentEvents</span></span><br><span data-ttu-id="c3d2f-164">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-164">
        - File</span></span><br><span data-ttu-id="c3d2f-165">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-165">
        - MatrixBindings</span></span><br><span data-ttu-id="c3d2f-166">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-166">
        - MatrixCoercion</span></span><br><span data-ttu-id="c3d2f-167">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="c3d2f-167">
        - Selection</span></span><br><span data-ttu-id="c3d2f-168">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="c3d2f-168">
        - Settings</span></span><br><span data-ttu-id="c3d2f-169">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-169">
        - TableBindings</span></span><br><span data-ttu-id="c3d2f-170">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-170">
        - TableCoercion</span></span><br><span data-ttu-id="c3d2f-171">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-171">
        - TextBindings</span></span><br><span data-ttu-id="c3d2f-172">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-172">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-173">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="c3d2f-173">Office 2019 on Windows</span></span><br><span data-ttu-id="c3d2f-174">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-174">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="c3d2f-175">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c3d2f-175">- TaskPane</span></span><br><span data-ttu-id="c3d2f-176">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-176">
        - Content</span></span><br><span data-ttu-id="c3d2f-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="c3d2f-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c3d2f-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c3d2f-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c3d2f-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c3d2f-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c3d2f-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c3d2f-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c3d2f-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c3d2f-188">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-188">- BindingEvents</span></span><br><span data-ttu-id="c3d2f-189">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-189">
        - CompressedFile</span></span><br><span data-ttu-id="c3d2f-190">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-190">
        - DocumentEvents</span></span><br><span data-ttu-id="c3d2f-191">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-191">
        - File</span></span><br><span data-ttu-id="c3d2f-192">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-192">
        - MatrixBindings</span></span><br><span data-ttu-id="c3d2f-193">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-193">
        - MatrixCoercion</span></span><br><span data-ttu-id="c3d2f-194">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="c3d2f-194">
        - Selection</span></span><br><span data-ttu-id="c3d2f-195">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="c3d2f-195">
        - Settings</span></span><br><span data-ttu-id="c3d2f-196">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-196">
        - TableBindings</span></span><br><span data-ttu-id="c3d2f-197">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-197">
        - TableCoercion</span></span><br><span data-ttu-id="c3d2f-198">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-198">
        - TextBindings</span></span><br><span data-ttu-id="c3d2f-199">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-199">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-200">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="c3d2f-200">Office 2016 on Windows</span></span><br><span data-ttu-id="c3d2f-201">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-201">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="c3d2f-202">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c3d2f-202">- TaskPane</span></span><br><span data-ttu-id="c3d2f-203">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-203">
        - Content</span></span></td>
    <td><span data-ttu-id="c3d2f-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="c3d2f-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="c3d2f-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c3d2f-207">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-207">- BindingEvents</span></span><br><span data-ttu-id="c3d2f-208">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-208">
        - CompressedFile</span></span><br><span data-ttu-id="c3d2f-209">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-209">
        - DocumentEvents</span></span><br><span data-ttu-id="c3d2f-210">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-210">
        - File</span></span><br><span data-ttu-id="c3d2f-211">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-211">
        - MatrixBindings</span></span><br><span data-ttu-id="c3d2f-212">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-212">
        - MatrixCoercion</span></span><br><span data-ttu-id="c3d2f-213">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="c3d2f-213">
        - Selection</span></span><br><span data-ttu-id="c3d2f-214">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="c3d2f-214">
        - Settings</span></span><br><span data-ttu-id="c3d2f-215">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-215">
        - TableBindings</span></span><br><span data-ttu-id="c3d2f-216">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-216">
        - TableCoercion</span></span><br><span data-ttu-id="c3d2f-217">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-217">
        - TextBindings</span></span><br><span data-ttu-id="c3d2f-218">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-218">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-219">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="c3d2f-219">Office 2013 on Windows</span></span><br><span data-ttu-id="c3d2f-220">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-220">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="c3d2f-221">
        - TaskPane</span><span class="sxs-lookup"><span data-stu-id="c3d2f-221">
        - TaskPane</span></span><br><span data-ttu-id="c3d2f-222">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-222">
        - Content</span></span></td>
    <td>  <span data-ttu-id="c3d2f-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c3d2f-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="c3d2f-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c3d2f-225">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-225">
        - BindingEvents</span></span><br><span data-ttu-id="c3d2f-226">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-226">
        - CompressedFile</span></span><br><span data-ttu-id="c3d2f-227">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-227">
        - DocumentEvents</span></span><br><span data-ttu-id="c3d2f-228">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-228">
        - File</span></span><br><span data-ttu-id="c3d2f-229">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-229">
        - MatrixBindings</span></span><br><span data-ttu-id="c3d2f-230">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-230">
        - MatrixCoercion</span></span><br><span data-ttu-id="c3d2f-231">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="c3d2f-231">
        - Selection</span></span><br><span data-ttu-id="c3d2f-232">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="c3d2f-232">
        - Settings</span></span><br><span data-ttu-id="c3d2f-233">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-233">
        - TableBindings</span></span><br><span data-ttu-id="c3d2f-234">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-234">
        - TableCoercion</span></span><br><span data-ttu-id="c3d2f-235">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-235">
        - TextBindings</span></span><br><span data-ttu-id="c3d2f-236">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-236">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-237">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="c3d2f-237">Office on iPad</span></span><br><span data-ttu-id="c3d2f-238">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-238">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="c3d2f-239">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c3d2f-239">- TaskPane</span></span><br><span data-ttu-id="c3d2f-240">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-240">
        - Content</span></span></td>
    <td><span data-ttu-id="c3d2f-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c3d2f-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c3d2f-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c3d2f-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c3d2f-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c3d2f-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c3d2f-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c3d2f-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="c3d2f-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="c3d2f-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c3d2f-253">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-253">- BindingEvents</span></span><br><span data-ttu-id="c3d2f-254">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-254">
        - DocumentEvents</span></span><br><span data-ttu-id="c3d2f-255">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-255">
        - File</span></span><br><span data-ttu-id="c3d2f-256">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-256">
        - MatrixBindings</span></span><br><span data-ttu-id="c3d2f-257">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-257">
        - MatrixCoercion</span></span><br><span data-ttu-id="c3d2f-258">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="c3d2f-258">
        - Selection</span></span><br><span data-ttu-id="c3d2f-259">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="c3d2f-259">
        - Settings</span></span><br><span data-ttu-id="c3d2f-260">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-260">
        - TableBindings</span></span><br><span data-ttu-id="c3d2f-261">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-261">
        - TableCoercion</span></span><br><span data-ttu-id="c3d2f-262">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-262">
        - TextBindings</span></span><br><span data-ttu-id="c3d2f-263">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-263">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-264">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="c3d2f-264">Office on Mac</span></span><br><span data-ttu-id="c3d2f-265">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-265">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="c3d2f-266">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c3d2f-266">- TaskPane</span></span><br><span data-ttu-id="c3d2f-267">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-267">
        - Content</span></span><br><span data-ttu-id="c3d2f-268">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="c3d2f-268">
        - Custom Functions</span></span><br><span data-ttu-id="c3d2f-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="c3d2f-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c3d2f-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c3d2f-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c3d2f-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c3d2f-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c3d2f-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c3d2f-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c3d2f-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="c3d2f-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="c3d2f-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c3d2f-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="c3d2f-283">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-283">- BindingEvents</span></span><br><span data-ttu-id="c3d2f-284">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-284">
        - CompressedFile</span></span><br><span data-ttu-id="c3d2f-285">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-285">
        - DocumentEvents</span></span><br><span data-ttu-id="c3d2f-286">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-286">
        - File</span></span><br><span data-ttu-id="c3d2f-287">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-287">
        - MatrixBindings</span></span><br><span data-ttu-id="c3d2f-288">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-288">
        - MatrixCoercion</span></span><br><span data-ttu-id="c3d2f-289">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-289">
        - PdfFile</span></span><br><span data-ttu-id="c3d2f-290">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="c3d2f-290">
        - Selection</span></span><br><span data-ttu-id="c3d2f-291">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="c3d2f-291">
        - Settings</span></span><br><span data-ttu-id="c3d2f-292">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-292">
        - TableBindings</span></span><br><span data-ttu-id="c3d2f-293">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-293">
        - TableCoercion</span></span><br><span data-ttu-id="c3d2f-294">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-294">
        - TextBindings</span></span><br><span data-ttu-id="c3d2f-295">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-295">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-296">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="c3d2f-296">Office 2019 on Mac</span></span><br><span data-ttu-id="c3d2f-297">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-297">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="c3d2f-298">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c3d2f-298">- TaskPane</span></span><br><span data-ttu-id="c3d2f-299">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-299">
        - Content</span></span><br><span data-ttu-id="c3d2f-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="c3d2f-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c3d2f-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c3d2f-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c3d2f-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c3d2f-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c3d2f-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c3d2f-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c3d2f-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c3d2f-311">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-311">- BindingEvents</span></span><br><span data-ttu-id="c3d2f-312">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-312">
        - CompressedFile</span></span><br><span data-ttu-id="c3d2f-313">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-313">
        - DocumentEvents</span></span><br><span data-ttu-id="c3d2f-314">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-314">
        - File</span></span><br><span data-ttu-id="c3d2f-315">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-315">
        - MatrixBindings</span></span><br><span data-ttu-id="c3d2f-316">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-316">
        - MatrixCoercion</span></span><br><span data-ttu-id="c3d2f-317">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-317">
        - PdfFile</span></span><br><span data-ttu-id="c3d2f-318">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="c3d2f-318">
        - Selection</span></span><br><span data-ttu-id="c3d2f-319">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="c3d2f-319">
        - Settings</span></span><br><span data-ttu-id="c3d2f-320">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-320">
        - TableBindings</span></span><br><span data-ttu-id="c3d2f-321">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-321">
        - TableCoercion</span></span><br><span data-ttu-id="c3d2f-322">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-322">
        - TextBindings</span></span><br><span data-ttu-id="c3d2f-323">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-323">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-324">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="c3d2f-324">Office 2016 on Mac</span></span><br><span data-ttu-id="c3d2f-325">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-325">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="c3d2f-326">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c3d2f-326">- TaskPane</span></span><br><span data-ttu-id="c3d2f-327">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-327">
        - Content</span></span></td>
    <td><span data-ttu-id="c3d2f-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="c3d2f-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="c3d2f-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c3d2f-331">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-331">- BindingEvents</span></span><br><span data-ttu-id="c3d2f-332">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-332">
        - CompressedFile</span></span><br><span data-ttu-id="c3d2f-333">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-333">
        - DocumentEvents</span></span><br><span data-ttu-id="c3d2f-334">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-334">
        - File</span></span><br><span data-ttu-id="c3d2f-335">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-335">
        - MatrixBindings</span></span><br><span data-ttu-id="c3d2f-336">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-336">
        - MatrixCoercion</span></span><br><span data-ttu-id="c3d2f-337">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-337">
        - PdfFile</span></span><br><span data-ttu-id="c3d2f-338">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="c3d2f-338">
        - Selection</span></span><br><span data-ttu-id="c3d2f-339">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="c3d2f-339">
        - Settings</span></span><br><span data-ttu-id="c3d2f-340">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-340">
        - TableBindings</span></span><br><span data-ttu-id="c3d2f-341">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-341">
        - TableCoercion</span></span><br><span data-ttu-id="c3d2f-342">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-342">
        - TextBindings</span></span><br><span data-ttu-id="c3d2f-343">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-343">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="c3d2f-344">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="c3d2f-344">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="c3d2f-345">Funções personalizadas (somente Excel)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-345">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="c3d2f-346">Plataforma</span><span class="sxs-lookup"><span data-stu-id="c3d2f-346">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="c3d2f-347">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="c3d2f-347">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="c3d2f-348">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="c3d2f-348">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="c3d2f-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-350">Office na Web</span><span class="sxs-lookup"><span data-stu-id="c3d2f-350">Office on the web</span></span></td>
    <td><span data-ttu-id="c3d2f-351">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="c3d2f-351">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="c3d2f-352">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-352">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-353">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="c3d2f-353">Office on Windows</span></span><br><span data-ttu-id="c3d2f-354">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-354">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="c3d2f-355">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="c3d2f-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="c3d2f-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-357">Office para Mac</span><span class="sxs-lookup"><span data-stu-id="c3d2f-357">Office for Mac</span></span><br><span data-ttu-id="c3d2f-358">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-358">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="c3d2f-359">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="c3d2f-359">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="c3d2f-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="c3d2f-361">Outlook</span><span class="sxs-lookup"><span data-stu-id="c3d2f-361">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c3d2f-362">Plataforma</span><span class="sxs-lookup"><span data-stu-id="c3d2f-362">Platform</span></span></th>
    <th><span data-ttu-id="c3d2f-363">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="c3d2f-363">Extension points</span></span></th>
    <th><span data-ttu-id="c3d2f-364">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="c3d2f-364">API requirement sets</span></span></th>
    <th><span data-ttu-id="c3d2f-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-366">Office na Web</span><span class="sxs-lookup"><span data-stu-id="c3d2f-366">Office on the web</span></span><br><span data-ttu-id="c3d2f-367">(moderno)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-367">(modern)</span></span></td>
    <td> <span data-ttu-id="c3d2f-368">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="c3d2f-368">- Mail Read</span></span><br><span data-ttu-id="c3d2f-369">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="c3d2f-369">
      - Mail Compose</span></span><br><span data-ttu-id="c3d2f-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c3d2f-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c3d2f-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c3d2f-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c3d2f-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c3d2f-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c3d2f-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="c3d2f-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="c3d2f-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="c3d2f-379">Não disponível</span><span class="sxs-lookup"><span data-stu-id="c3d2f-379">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-380">Office na Web</span><span class="sxs-lookup"><span data-stu-id="c3d2f-380">Office on the web</span></span><br><span data-ttu-id="c3d2f-381">(clássico)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-381">(classic)</span></span></td>
    <td> <span data-ttu-id="c3d2f-382">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="c3d2f-382">- Mail Read</span></span><br><span data-ttu-id="c3d2f-383">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="c3d2f-383">
      - Mail Compose</span></span><br><span data-ttu-id="c3d2f-384">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-384">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c3d2f-385">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-385">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c3d2f-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c3d2f-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c3d2f-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c3d2f-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c3d2f-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="c3d2f-391">Não disponível</span><span class="sxs-lookup"><span data-stu-id="c3d2f-391">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-392">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="c3d2f-392">Office on Windows</span></span><br><span data-ttu-id="c3d2f-393">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-393">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c3d2f-394">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="c3d2f-394">- Mail Read</span></span><br><span data-ttu-id="c3d2f-395">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="c3d2f-395">
      - Mail Compose</span></span><br><span data-ttu-id="c3d2f-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="c3d2f-397">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="c3d2f-397">
      - Modules</span></span></td>
    <td> <span data-ttu-id="c3d2f-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c3d2f-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c3d2f-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c3d2f-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c3d2f-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c3d2f-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="c3d2f-404">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-404">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="c3d2f-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="c3d2f-406">Não disponível</span><span class="sxs-lookup"><span data-stu-id="c3d2f-406">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-407">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="c3d2f-407">Office 2019 on Windows</span></span><br><span data-ttu-id="c3d2f-408">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-408">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c3d2f-409">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="c3d2f-409">- Mail Read</span></span><br><span data-ttu-id="c3d2f-410">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="c3d2f-410">
      - Mail Compose</span></span><br><span data-ttu-id="c3d2f-411">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-411">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="c3d2f-412">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="c3d2f-412">
      - Modules</span></span></td>
    <td> <span data-ttu-id="c3d2f-413">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-413">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c3d2f-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c3d2f-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c3d2f-416">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-416">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c3d2f-417">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-417">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c3d2f-418">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-418">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="c3d2f-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="c3d2f-420">Não disponível</span><span class="sxs-lookup"><span data-stu-id="c3d2f-420">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-421">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="c3d2f-421">Office 2016 on Windows</span></span><br><span data-ttu-id="c3d2f-422">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-422">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c3d2f-423">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="c3d2f-423">- Mail Read</span></span><br><span data-ttu-id="c3d2f-424">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="c3d2f-424">
      - Mail Compose</span></span><br><span data-ttu-id="c3d2f-425">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-425">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="c3d2f-426">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="c3d2f-426">
      - Modules</span></span></td>
    <td> <span data-ttu-id="c3d2f-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c3d2f-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c3d2f-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c3d2f-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="c3d2f-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="c3d2f-431">Não disponível</span><span class="sxs-lookup"><span data-stu-id="c3d2f-431">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-432">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="c3d2f-432">Office 2013 on Windows</span></span><br><span data-ttu-id="c3d2f-433">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-433">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c3d2f-434">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="c3d2f-434">- Mail Read</span></span><br><span data-ttu-id="c3d2f-435">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="c3d2f-435">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="c3d2f-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c3d2f-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c3d2f-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="c3d2f-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="c3d2f-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="c3d2f-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="c3d2f-440">Não disponível</span><span class="sxs-lookup"><span data-stu-id="c3d2f-440">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-441">Office no iOS</span><span class="sxs-lookup"><span data-stu-id="c3d2f-441">Office on iOS</span></span><br><span data-ttu-id="c3d2f-442">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-442">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c3d2f-443">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="c3d2f-443">- Mail Read</span></span><br><span data-ttu-id="c3d2f-444">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-444">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c3d2f-445">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-445">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c3d2f-446">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-446">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c3d2f-447">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-447">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c3d2f-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c3d2f-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="c3d2f-450">Não disponível</span><span class="sxs-lookup"><span data-stu-id="c3d2f-450">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-451">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="c3d2f-451">Office on Mac</span></span><br><span data-ttu-id="c3d2f-452">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-452">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c3d2f-453">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="c3d2f-453">- Mail Read</span></span><br><span data-ttu-id="c3d2f-454">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="c3d2f-454">
      - Mail Compose</span></span><br><span data-ttu-id="c3d2f-455">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-455">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c3d2f-456">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-456">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c3d2f-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c3d2f-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c3d2f-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c3d2f-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c3d2f-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="c3d2f-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="c3d2f-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="c3d2f-464">Não disponível</span><span class="sxs-lookup"><span data-stu-id="c3d2f-464">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-465">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="c3d2f-465">Office 2019 on Mac</span></span><br><span data-ttu-id="c3d2f-466">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-466">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c3d2f-467">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="c3d2f-467">- Mail Read</span></span><br><span data-ttu-id="c3d2f-468">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="c3d2f-468">
      - Mail Compose</span></span><br><span data-ttu-id="c3d2f-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c3d2f-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c3d2f-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c3d2f-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c3d2f-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c3d2f-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c3d2f-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="c3d2f-476">Não disponível</span><span class="sxs-lookup"><span data-stu-id="c3d2f-476">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-477">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="c3d2f-477">Office 2016 on Mac</span></span><br><span data-ttu-id="c3d2f-478">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-478">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c3d2f-479">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="c3d2f-479">- Mail Read</span></span><br><span data-ttu-id="c3d2f-480">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="c3d2f-480">
      - Mail Compose</span></span><br><span data-ttu-id="c3d2f-481">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-481">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c3d2f-482">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-482">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c3d2f-483">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-483">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c3d2f-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c3d2f-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c3d2f-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c3d2f-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="c3d2f-488">Não disponível</span><span class="sxs-lookup"><span data-stu-id="c3d2f-488">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-489">Outlook no Android</span><span class="sxs-lookup"><span data-stu-id="c3d2f-489">Office on Android</span></span><br><span data-ttu-id="c3d2f-490">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-490">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c3d2f-491">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="c3d2f-491">- Mail Read</span></span><br><span data-ttu-id="c3d2f-492">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-492">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c3d2f-493">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-493">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c3d2f-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c3d2f-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c3d2f-496">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-496">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c3d2f-497">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-497">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="c3d2f-498">Não disponível</span><span class="sxs-lookup"><span data-stu-id="c3d2f-498">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="c3d2f-499">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="c3d2f-499">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c3d2f-500">O suporte ao cliente para um conjunto de requisitos pode ser restringido pelo suporte do servidor Exchange.</span><span class="sxs-lookup"><span data-stu-id="c3d2f-500">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="c3d2f-501">Consulte [Conjuntos de requisitos da API JavaScript do Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) para obter detalhes sobre o intervalo de conjuntos de requisitos suportado pelo servidor Exchange e pelos clientes Outlook.</span><span class="sxs-lookup"><span data-stu-id="c3d2f-501">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="c3d2f-502">Word</span><span class="sxs-lookup"><span data-stu-id="c3d2f-502">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c3d2f-503">Plataforma</span><span class="sxs-lookup"><span data-stu-id="c3d2f-503">Platform</span></span></th>
    <th><span data-ttu-id="c3d2f-504">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="c3d2f-504">Extension points</span></span></th>
    <th><span data-ttu-id="c3d2f-505">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="c3d2f-505">API requirement sets</span></span></th>
    <th><span data-ttu-id="c3d2f-506"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-506"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-507">Office na Web</span><span class="sxs-lookup"><span data-stu-id="c3d2f-507">Office on the web</span></span></td>
    <td> <span data-ttu-id="c3d2f-508">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c3d2f-508">- TaskPane</span></span><br><span data-ttu-id="c3d2f-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c3d2f-510">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-510">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-511">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-511">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c3d2f-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c3d2f-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-514">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-514">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c3d2f-515">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-515">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="c3d2f-516">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-516">- BindingEvents</span></span><br><span data-ttu-id="c3d2f-517">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c3d2f-517">
         - CustomXmlParts</span></span><br><span data-ttu-id="c3d2f-518">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-518">
         - DocumentEvents</span></span><br><span data-ttu-id="c3d2f-519">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-519">
         - File</span></span><br><span data-ttu-id="c3d2f-520">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-520">
         - HtmlCoercion</span></span><br><span data-ttu-id="c3d2f-521">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-521">
         - MatrixBindings</span></span><br><span data-ttu-id="c3d2f-522">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-522">
         - MatrixCoercion</span></span><br><span data-ttu-id="c3d2f-523">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-523">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c3d2f-524">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-524">
         - PdfFile</span></span><br><span data-ttu-id="c3d2f-525">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="c3d2f-525">
         - Selection</span></span><br><span data-ttu-id="c3d2f-526">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="c3d2f-526">
         - Settings</span></span><br><span data-ttu-id="c3d2f-527">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-527">
         - TableBindings</span></span><br><span data-ttu-id="c3d2f-528">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-528">
         - TableCoercion</span></span><br><span data-ttu-id="c3d2f-529">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-529">
         - TextBindings</span></span><br><span data-ttu-id="c3d2f-530">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-530">
         - TextCoercion</span></span><br><span data-ttu-id="c3d2f-531">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-531">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-532">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="c3d2f-532">Office on Windows</span></span><br><span data-ttu-id="c3d2f-533">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-533">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c3d2f-534">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c3d2f-534">- TaskPane</span></span><br><span data-ttu-id="c3d2f-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c3d2f-536">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-536">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-537">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-537">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c3d2f-538">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-538">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c3d2f-539">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-539">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-540">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-540">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c3d2f-541">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-541">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="c3d2f-542">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-542">- BindingEvents</span></span><br><span data-ttu-id="c3d2f-543">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-543">
         - CompressedFile</span></span><br><span data-ttu-id="c3d2f-544">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c3d2f-544">
         - CustomXmlParts</span></span><br><span data-ttu-id="c3d2f-545">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-545">
         - DocumentEvents</span></span><br><span data-ttu-id="c3d2f-546">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-546">
         - File</span></span><br><span data-ttu-id="c3d2f-547">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-547">
         - HtmlCoercion</span></span><br><span data-ttu-id="c3d2f-548">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-548">
         - MatrixBindings</span></span><br><span data-ttu-id="c3d2f-549">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-549">
         - MatrixCoercion</span></span><br><span data-ttu-id="c3d2f-550">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-550">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c3d2f-551">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-551">
         - PdfFile</span></span><br><span data-ttu-id="c3d2f-552">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="c3d2f-552">
         - Selection</span></span><br><span data-ttu-id="c3d2f-553">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="c3d2f-553">
         - Settings</span></span><br><span data-ttu-id="c3d2f-554">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-554">
         - TableBindings</span></span><br><span data-ttu-id="c3d2f-555">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-555">
         - TableCoercion</span></span><br><span data-ttu-id="c3d2f-556">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-556">
         - TextBindings</span></span><br><span data-ttu-id="c3d2f-557">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-557">
         - TextCoercion</span></span><br><span data-ttu-id="c3d2f-558">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-558">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-559">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="c3d2f-559">Office 2019 on Windows</span></span><br><span data-ttu-id="c3d2f-560">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-560">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c3d2f-561">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c3d2f-561">- TaskPane</span></span><br><span data-ttu-id="c3d2f-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c3d2f-563">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-563">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-564">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-564">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c3d2f-565">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-565">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c3d2f-566">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-566">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-567">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-567">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c3d2f-568">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-568">- BindingEvents</span></span><br><span data-ttu-id="c3d2f-569">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-569">
         - CompressedFile</span></span><br><span data-ttu-id="c3d2f-570">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c3d2f-570">
         - CustomXmlParts</span></span><br><span data-ttu-id="c3d2f-571">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-571">
         - DocumentEvents</span></span><br><span data-ttu-id="c3d2f-572">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-572">
         - File</span></span><br><span data-ttu-id="c3d2f-573">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-573">
         - HtmlCoercion</span></span><br><span data-ttu-id="c3d2f-574">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-574">
         - MatrixBindings</span></span><br><span data-ttu-id="c3d2f-575">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-575">
         - MatrixCoercion</span></span><br><span data-ttu-id="c3d2f-576">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-576">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c3d2f-577">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-577">
         - PdfFile</span></span><br><span data-ttu-id="c3d2f-578">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="c3d2f-578">
         - Selection</span></span><br><span data-ttu-id="c3d2f-579">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="c3d2f-579">
         - Settings</span></span><br><span data-ttu-id="c3d2f-580">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-580">
         - TableBindings</span></span><br><span data-ttu-id="c3d2f-581">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-581">
         - TableCoercion</span></span><br><span data-ttu-id="c3d2f-582">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-582">
         - TextBindings</span></span><br><span data-ttu-id="c3d2f-583">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-583">
         - TextCoercion</span></span><br><span data-ttu-id="c3d2f-584">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-584">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-585">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="c3d2f-585">Office 2016 on Windows</span></span><br><span data-ttu-id="c3d2f-586">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-586">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c3d2f-587">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c3d2f-587">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c3d2f-588">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-588">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-589">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="c3d2f-589">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="c3d2f-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c3d2f-591">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-591">- BindingEvents</span></span><br><span data-ttu-id="c3d2f-592">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-592">
         - CompressedFile</span></span><br><span data-ttu-id="c3d2f-593">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c3d2f-593">
         - CustomXmlParts</span></span><br><span data-ttu-id="c3d2f-594">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-594">
         - DocumentEvents</span></span><br><span data-ttu-id="c3d2f-595">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-595">
         - File</span></span><br><span data-ttu-id="c3d2f-596">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-596">
         - HtmlCoercion</span></span><br><span data-ttu-id="c3d2f-597">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-597">
         - MatrixBindings</span></span><br><span data-ttu-id="c3d2f-598">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-598">
         - MatrixCoercion</span></span><br><span data-ttu-id="c3d2f-599">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-599">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c3d2f-600">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-600">
         - PdfFile</span></span><br><span data-ttu-id="c3d2f-601">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="c3d2f-601">
         - Selection</span></span><br><span data-ttu-id="c3d2f-602">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="c3d2f-602">
         - Settings</span></span><br><span data-ttu-id="c3d2f-603">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-603">
         - TableBindings</span></span><br><span data-ttu-id="c3d2f-604">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-604">
         - TableCoercion</span></span><br><span data-ttu-id="c3d2f-605">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-605">
         - TextBindings</span></span><br><span data-ttu-id="c3d2f-606">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-606">
         - TextCoercion</span></span><br><span data-ttu-id="c3d2f-607">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-607">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-608">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="c3d2f-608">Office 2013 on Windows</span></span><br><span data-ttu-id="c3d2f-609">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-609">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c3d2f-610">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c3d2f-610">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c3d2f-611">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c3d2f-611">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="c3d2f-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c3d2f-613">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-613">- BindingEvents</span></span><br><span data-ttu-id="c3d2f-614">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-614">
         - CompressedFile</span></span><br><span data-ttu-id="c3d2f-615">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c3d2f-615">
         - CustomXmlParts</span></span><br><span data-ttu-id="c3d2f-616">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-616">
         - DocumentEvents</span></span><br><span data-ttu-id="c3d2f-617">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-617">
         - File</span></span><br><span data-ttu-id="c3d2f-618">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-618">
         - HtmlCoercion</span></span><br><span data-ttu-id="c3d2f-619">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-619">
         - MatrixBindings</span></span><br><span data-ttu-id="c3d2f-620">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-620">
         - MatrixCoercion</span></span><br><span data-ttu-id="c3d2f-621">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-621">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c3d2f-622">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-622">
         - PdfFile</span></span><br><span data-ttu-id="c3d2f-623">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="c3d2f-623">
         - Selection</span></span><br><span data-ttu-id="c3d2f-624">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="c3d2f-624">
         - Settings</span></span><br><span data-ttu-id="c3d2f-625">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-625">
         - TableBindings</span></span><br><span data-ttu-id="c3d2f-626">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-626">
         - TableCoercion</span></span><br><span data-ttu-id="c3d2f-627">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-627">
         - TextBindings</span></span><br><span data-ttu-id="c3d2f-628">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-628">
         - TextCoercion</span></span><br><span data-ttu-id="c3d2f-629">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-629">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-630">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="c3d2f-630">Office on iPad</span></span><br><span data-ttu-id="c3d2f-631">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-631">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c3d2f-632">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c3d2f-632">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c3d2f-633">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-633">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c3d2f-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c3d2f-636">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-636">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-637">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-637">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="c3d2f-638">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-638">- BindingEvents</span></span><br><span data-ttu-id="c3d2f-639">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-639">
         - CompressedFile</span></span><br><span data-ttu-id="c3d2f-640">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c3d2f-640">
         - CustomXmlParts</span></span><br><span data-ttu-id="c3d2f-641">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-641">
         - DocumentEvents</span></span><br><span data-ttu-id="c3d2f-642">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-642">
         - File</span></span><br><span data-ttu-id="c3d2f-643">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-643">
         - HtmlCoercion</span></span><br><span data-ttu-id="c3d2f-644">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-644">
         - MatrixBindings</span></span><br><span data-ttu-id="c3d2f-645">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-645">
         - MatrixCoercion</span></span><br><span data-ttu-id="c3d2f-646">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-646">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c3d2f-647">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-647">
         - PdfFile</span></span><br><span data-ttu-id="c3d2f-648">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="c3d2f-648">
         - Selection</span></span><br><span data-ttu-id="c3d2f-649">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="c3d2f-649">
         - Settings</span></span><br><span data-ttu-id="c3d2f-650">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-650">
         - TableBindings</span></span><br><span data-ttu-id="c3d2f-651">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-651">
         - TableCoercion</span></span><br><span data-ttu-id="c3d2f-652">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-652">
         - TextBindings</span></span><br><span data-ttu-id="c3d2f-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-653">
         - TextCoercion</span></span><br><span data-ttu-id="c3d2f-654">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-654">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-655">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="c3d2f-655">Office on Mac</span></span><br><span data-ttu-id="c3d2f-656">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-656">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c3d2f-657">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c3d2f-657">- TaskPane</span></span><br><span data-ttu-id="c3d2f-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c3d2f-659">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-659">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-660">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-660">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c3d2f-661">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-661">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c3d2f-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c3d2f-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="c3d2f-665">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-665">- BindingEvents</span></span><br><span data-ttu-id="c3d2f-666">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-666">
         - CompressedFile</span></span><br><span data-ttu-id="c3d2f-667">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c3d2f-667">
         - CustomXmlParts</span></span><br><span data-ttu-id="c3d2f-668">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-668">
         - DocumentEvents</span></span><br><span data-ttu-id="c3d2f-669">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-669">
         - File</span></span><br><span data-ttu-id="c3d2f-670">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-670">
         - HtmlCoercion</span></span><br><span data-ttu-id="c3d2f-671">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-671">
         - MatrixBindings</span></span><br><span data-ttu-id="c3d2f-672">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-672">
         - MatrixCoercion</span></span><br><span data-ttu-id="c3d2f-673">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-673">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c3d2f-674">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-674">
         - PdfFile</span></span><br><span data-ttu-id="c3d2f-675">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="c3d2f-675">
         - Selection</span></span><br><span data-ttu-id="c3d2f-676">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="c3d2f-676">
         - Settings</span></span><br><span data-ttu-id="c3d2f-677">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-677">
         - TableBindings</span></span><br><span data-ttu-id="c3d2f-678">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-678">
         - TableCoercion</span></span><br><span data-ttu-id="c3d2f-679">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-679">
         - TextBindings</span></span><br><span data-ttu-id="c3d2f-680">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-680">
         - TextCoercion</span></span><br><span data-ttu-id="c3d2f-681">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-681">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-682">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="c3d2f-682">Office 2019 on Mac</span></span><br><span data-ttu-id="c3d2f-683">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-683">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c3d2f-684">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c3d2f-684">- TaskPane</span></span><br><span data-ttu-id="c3d2f-685">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-685">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c3d2f-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c3d2f-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c3d2f-689">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-689">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-690">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-690">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="c3d2f-691">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-691">- BindingEvents</span></span><br><span data-ttu-id="c3d2f-692">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-692">
         - CompressedFile</span></span><br><span data-ttu-id="c3d2f-693">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c3d2f-693">
         - CustomXmlParts</span></span><br><span data-ttu-id="c3d2f-694">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-694">
         - DocumentEvents</span></span><br><span data-ttu-id="c3d2f-695">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-695">
         - File</span></span><br><span data-ttu-id="c3d2f-696">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-696">
         - HtmlCoercion</span></span><br><span data-ttu-id="c3d2f-697">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-697">
         - MatrixBindings</span></span><br><span data-ttu-id="c3d2f-698">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-698">
         - MatrixCoercion</span></span><br><span data-ttu-id="c3d2f-699">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-699">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c3d2f-700">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-700">
         - PdfFile</span></span><br><span data-ttu-id="c3d2f-701">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="c3d2f-701">
         - Selection</span></span><br><span data-ttu-id="c3d2f-702">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="c3d2f-702">
         - Settings</span></span><br><span data-ttu-id="c3d2f-703">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-703">
         - TableBindings</span></span><br><span data-ttu-id="c3d2f-704">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-704">
         - TableCoercion</span></span><br><span data-ttu-id="c3d2f-705">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-705">
         - TextBindings</span></span><br><span data-ttu-id="c3d2f-706">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-706">
         - TextCoercion</span></span><br><span data-ttu-id="c3d2f-707">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-707">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-708">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="c3d2f-708">Office 2016 on Mac</span></span><br><span data-ttu-id="c3d2f-709">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-709">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c3d2f-710">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c3d2f-710">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c3d2f-711">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-711">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="c3d2f-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="c3d2f-713">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-713">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c3d2f-714">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-714">- BindingEvents</span></span><br><span data-ttu-id="c3d2f-715">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-715">
         - CompressedFile</span></span><br><span data-ttu-id="c3d2f-716">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c3d2f-716">
         - CustomXmlParts</span></span><br><span data-ttu-id="c3d2f-717">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-717">
         - DocumentEvents</span></span><br><span data-ttu-id="c3d2f-718">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-718">
         - File</span></span><br><span data-ttu-id="c3d2f-719">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-719">
         - HtmlCoercion</span></span><br><span data-ttu-id="c3d2f-720">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-720">
         - MatrixBindings</span></span><br><span data-ttu-id="c3d2f-721">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-721">
         - MatrixCoercion</span></span><br><span data-ttu-id="c3d2f-722">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-722">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c3d2f-723">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-723">
         - PdfFile</span></span><br><span data-ttu-id="c3d2f-724">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="c3d2f-724">
         - Selection</span></span><br><span data-ttu-id="c3d2f-725">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="c3d2f-725">
         - Settings</span></span><br><span data-ttu-id="c3d2f-726">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-726">
         - TableBindings</span></span><br><span data-ttu-id="c3d2f-727">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-727">
         - TableCoercion</span></span><br><span data-ttu-id="c3d2f-728">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c3d2f-728">
         - TextBindings</span></span><br><span data-ttu-id="c3d2f-729">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-729">
         - TextCoercion</span></span><br><span data-ttu-id="c3d2f-730">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-730">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="c3d2f-731">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="c3d2f-731">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="c3d2f-732">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="c3d2f-732">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c3d2f-733">Plataforma</span><span class="sxs-lookup"><span data-stu-id="c3d2f-733">Platform</span></span></th>
    <th><span data-ttu-id="c3d2f-734">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="c3d2f-734">Extension points</span></span></th>
    <th><span data-ttu-id="c3d2f-735">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="c3d2f-735">API requirement sets</span></span></th>
    <th><span data-ttu-id="c3d2f-736"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-736"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-737">Office na Web</span><span class="sxs-lookup"><span data-stu-id="c3d2f-737">Office on the web</span></span></td>
    <td> <span data-ttu-id="c3d2f-738">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-738">- Content</span></span><br><span data-ttu-id="c3d2f-739">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="c3d2f-739">
         - TaskPane</span></span><br><span data-ttu-id="c3d2f-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c3d2f-741">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-741">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-742">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-742">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-743">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-743">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c3d2f-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="c3d2f-745">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c3d2f-745">- ActiveView</span></span><br><span data-ttu-id="c3d2f-746">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-746">
         - CompressedFile</span></span><br><span data-ttu-id="c3d2f-747">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-747">
         - DocumentEvents</span></span><br><span data-ttu-id="c3d2f-748">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-748">
         - File</span></span><br><span data-ttu-id="c3d2f-749">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-749">
         - PdfFile</span></span><br><span data-ttu-id="c3d2f-750">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="c3d2f-750">
         - Selection</span></span><br><span data-ttu-id="c3d2f-751">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="c3d2f-751">
         - Settings</span></span><br><span data-ttu-id="c3d2f-752">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-752">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-753">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="c3d2f-753">Office on Windows</span></span><br><span data-ttu-id="c3d2f-754">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-754">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c3d2f-755">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-755">- Content</span></span><br><span data-ttu-id="c3d2f-756">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="c3d2f-756">
         - TaskPane</span></span><br><span data-ttu-id="c3d2f-757">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-757">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c3d2f-758">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-758">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c3d2f-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="c3d2f-762">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c3d2f-762">- ActiveView</span></span><br><span data-ttu-id="c3d2f-763">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-763">
         - CompressedFile</span></span><br><span data-ttu-id="c3d2f-764">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-764">
         - DocumentEvents</span></span><br><span data-ttu-id="c3d2f-765">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-765">
         - File</span></span><br><span data-ttu-id="c3d2f-766">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-766">
         - PdfFile</span></span><br><span data-ttu-id="c3d2f-767">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="c3d2f-767">
         - Selection</span></span><br><span data-ttu-id="c3d2f-768">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="c3d2f-768">
         - Settings</span></span><br><span data-ttu-id="c3d2f-769">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-769">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-770">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="c3d2f-770">Office 2019 on Windows</span></span><br><span data-ttu-id="c3d2f-771">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-771">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c3d2f-772">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-772">- Content</span></span><br><span data-ttu-id="c3d2f-773">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="c3d2f-773">
         - TaskPane</span></span><br><span data-ttu-id="c3d2f-774">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-774">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c3d2f-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c3d2f-777">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c3d2f-777">- ActiveView</span></span><br><span data-ttu-id="c3d2f-778">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-778">
         - CompressedFile</span></span><br><span data-ttu-id="c3d2f-779">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-779">
         - DocumentEvents</span></span><br><span data-ttu-id="c3d2f-780">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-780">
         - File</span></span><br><span data-ttu-id="c3d2f-781">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-781">
         - PdfFile</span></span><br><span data-ttu-id="c3d2f-782">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="c3d2f-782">
         - Selection</span></span><br><span data-ttu-id="c3d2f-783">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="c3d2f-783">
         - Settings</span></span><br><span data-ttu-id="c3d2f-784">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-784">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-785">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="c3d2f-785">Office 2016 on Windows</span></span><br><span data-ttu-id="c3d2f-786">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-786">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c3d2f-787">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-787">- Content</span></span><br><span data-ttu-id="c3d2f-788">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="c3d2f-788">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="c3d2f-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c3d2f-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="c3d2f-790">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-790">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c3d2f-791">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c3d2f-791">- ActiveView</span></span><br><span data-ttu-id="c3d2f-792">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-792">
         - CompressedFile</span></span><br><span data-ttu-id="c3d2f-793">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-793">
         - DocumentEvents</span></span><br><span data-ttu-id="c3d2f-794">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-794">
         - File</span></span><br><span data-ttu-id="c3d2f-795">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-795">
         - PdfFile</span></span><br><span data-ttu-id="c3d2f-796">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="c3d2f-796">
         - Selection</span></span><br><span data-ttu-id="c3d2f-797">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="c3d2f-797">
         - Settings</span></span><br><span data-ttu-id="c3d2f-798">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-798">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-799">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="c3d2f-799">Office 2013 on Windows</span></span><br><span data-ttu-id="c3d2f-800">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-800">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c3d2f-801">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-801">- Content</span></span><br><span data-ttu-id="c3d2f-802">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="c3d2f-802">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="c3d2f-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c3d2f-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="c3d2f-804">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-804">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c3d2f-805">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c3d2f-805">- ActiveView</span></span><br><span data-ttu-id="c3d2f-806">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-806">
         - CompressedFile</span></span><br><span data-ttu-id="c3d2f-807">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-807">
         - DocumentEvents</span></span><br><span data-ttu-id="c3d2f-808">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-808">
         - File</span></span><br><span data-ttu-id="c3d2f-809">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-809">
         - PdfFile</span></span><br><span data-ttu-id="c3d2f-810">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="c3d2f-810">
         - Selection</span></span><br><span data-ttu-id="c3d2f-811">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="c3d2f-811">
         - Settings</span></span><br><span data-ttu-id="c3d2f-812">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-812">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-813">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="c3d2f-813">Office on iPad</span></span><br><span data-ttu-id="c3d2f-814">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-814">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c3d2f-815">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-815">- Content</span></span><br><span data-ttu-id="c3d2f-816">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="c3d2f-816">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="c3d2f-817">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-817">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-819">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-819">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c3d2f-820">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c3d2f-820">- ActiveView</span></span><br><span data-ttu-id="c3d2f-821">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-821">
         - CompressedFile</span></span><br><span data-ttu-id="c3d2f-822">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-822">
         - DocumentEvents</span></span><br><span data-ttu-id="c3d2f-823">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-823">
         - File</span></span><br><span data-ttu-id="c3d2f-824">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-824">
         - PdfFile</span></span><br><span data-ttu-id="c3d2f-825">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="c3d2f-825">
         - Selection</span></span><br><span data-ttu-id="c3d2f-826">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="c3d2f-826">
         - Settings</span></span><br><span data-ttu-id="c3d2f-827">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-827">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-828">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="c3d2f-828">Office on Mac</span></span><br><span data-ttu-id="c3d2f-829">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-829">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c3d2f-830">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-830">- Content</span></span><br><span data-ttu-id="c3d2f-831">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="c3d2f-831">
         - TaskPane</span></span><br><span data-ttu-id="c3d2f-832">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-832">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c3d2f-833">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-833">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-834">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-834">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-835">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-835">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c3d2f-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="c3d2f-837">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c3d2f-837">- ActiveView</span></span><br><span data-ttu-id="c3d2f-838">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-838">
         - CompressedFile</span></span><br><span data-ttu-id="c3d2f-839">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-839">
         - DocumentEvents</span></span><br><span data-ttu-id="c3d2f-840">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-840">
         - File</span></span><br><span data-ttu-id="c3d2f-841">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-841">
         - PdfFile</span></span><br><span data-ttu-id="c3d2f-842">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="c3d2f-842">
         - Selection</span></span><br><span data-ttu-id="c3d2f-843">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="c3d2f-843">
         - Settings</span></span><br><span data-ttu-id="c3d2f-844">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-844">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-845">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="c3d2f-845">Office 2019 on Mac</span></span><br><span data-ttu-id="c3d2f-846">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-846">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c3d2f-847">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-847">- Content</span></span><br><span data-ttu-id="c3d2f-848">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="c3d2f-848">
         - TaskPane</span></span><br><span data-ttu-id="c3d2f-849">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-849">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c3d2f-850">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-850">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c3d2f-852">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c3d2f-852">- ActiveView</span></span><br><span data-ttu-id="c3d2f-853">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-853">
         - CompressedFile</span></span><br><span data-ttu-id="c3d2f-854">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-854">
         - DocumentEvents</span></span><br><span data-ttu-id="c3d2f-855">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-855">
         - File</span></span><br><span data-ttu-id="c3d2f-856">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-856">
         - PdfFile</span></span><br><span data-ttu-id="c3d2f-857">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="c3d2f-857">
         - Selection</span></span><br><span data-ttu-id="c3d2f-858">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="c3d2f-858">
         - Settings</span></span><br><span data-ttu-id="c3d2f-859">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-859">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-860">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="c3d2f-860">Office 2016 on Mac</span></span><br><span data-ttu-id="c3d2f-861">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-861">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c3d2f-862">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-862">- Content</span></span><br><span data-ttu-id="c3d2f-863">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="c3d2f-863">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="c3d2f-864">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c3d2f-864">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="c3d2f-865">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-865">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c3d2f-866">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c3d2f-866">- ActiveView</span></span><br><span data-ttu-id="c3d2f-867">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-867">
         - CompressedFile</span></span><br><span data-ttu-id="c3d2f-868">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-868">
         - DocumentEvents</span></span><br><span data-ttu-id="c3d2f-869">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-869">
         - File</span></span><br><span data-ttu-id="c3d2f-870">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c3d2f-870">
         - PdfFile</span></span><br><span data-ttu-id="c3d2f-871">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="c3d2f-871">
         - Selection</span></span><br><span data-ttu-id="c3d2f-872">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="c3d2f-872">
         - Settings</span></span><br><span data-ttu-id="c3d2f-873">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-873">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="c3d2f-874">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="c3d2f-874">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="c3d2f-875">OneNote</span><span class="sxs-lookup"><span data-stu-id="c3d2f-875">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c3d2f-876">Plataforma</span><span class="sxs-lookup"><span data-stu-id="c3d2f-876">Platform</span></span></th>
    <th><span data-ttu-id="c3d2f-877">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="c3d2f-877">Extension points</span></span></th>
    <th><span data-ttu-id="c3d2f-878">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="c3d2f-878">API requirement sets</span></span></th>
    <th><span data-ttu-id="c3d2f-879"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-879"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-880">Office na Web</span><span class="sxs-lookup"><span data-stu-id="c3d2f-880">Office on the web</span></span></td>
    <td> <span data-ttu-id="c3d2f-881">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c3d2f-881">- Content</span></span><br><span data-ttu-id="c3d2f-882">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="c3d2f-882">
         - TaskPane</span></span><br><span data-ttu-id="c3d2f-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c3d2f-884">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-884">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-885">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-885">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c3d2f-886">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-886">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c3d2f-887">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d2f-887">- DocumentEvents</span></span><br><span data-ttu-id="c3d2f-888">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-888">
         - HtmlCoercion</span></span><br><span data-ttu-id="c3d2f-889">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="c3d2f-889">
         - Settings</span></span><br><span data-ttu-id="c3d2f-890">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-890">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="c3d2f-891">Project</span><span class="sxs-lookup"><span data-stu-id="c3d2f-891">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c3d2f-892">Plataforma</span><span class="sxs-lookup"><span data-stu-id="c3d2f-892">Platform</span></span></th>
    <th><span data-ttu-id="c3d2f-893">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="c3d2f-893">Extension points</span></span></th>
    <th><span data-ttu-id="c3d2f-894">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="c3d2f-894">API requirement sets</span></span></th>
    <th><span data-ttu-id="c3d2f-895"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-895"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-896">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="c3d2f-896">Office 2019 on Windows</span></span><br><span data-ttu-id="c3d2f-897">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-897">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c3d2f-898">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c3d2f-898">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c3d2f-899">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-899">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c3d2f-900">- Seleção</span><span class="sxs-lookup"><span data-stu-id="c3d2f-900">- Selection</span></span><br><span data-ttu-id="c3d2f-901">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-901">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-902">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="c3d2f-902">Office 2016 on Windows</span></span><br><span data-ttu-id="c3d2f-903">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-903">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c3d2f-904">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c3d2f-904">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c3d2f-905">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-905">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c3d2f-906">- Seleção</span><span class="sxs-lookup"><span data-stu-id="c3d2f-906">- Selection</span></span><br><span data-ttu-id="c3d2f-907">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-907">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d2f-908">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="c3d2f-908">Office 2013 on Windows</span></span><br><span data-ttu-id="c3d2f-909">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-909">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c3d2f-910">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c3d2f-910">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c3d2f-911">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d2f-911">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c3d2f-912">- Seleção</span><span class="sxs-lookup"><span data-stu-id="c3d2f-912">- Selection</span></span><br><span data-ttu-id="c3d2f-913">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d2f-913">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="c3d2f-914">Confira também</span><span class="sxs-lookup"><span data-stu-id="c3d2f-914">See also</span></span>

- [<span data-ttu-id="c3d2f-915">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="c3d2f-915">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="c3d2f-916">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="c3d2f-916">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="c3d2f-917">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="c3d2f-917">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="c3d2f-918">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="c3d2f-918">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="c3d2f-919">Documentação de Referência da API</span><span class="sxs-lookup"><span data-stu-id="c3d2f-919">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="c3d2f-920">Histórico de atualizações do Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="c3d2f-920">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="c3d2f-921">Histórico de atualizações do Office 2016 e 2019 (Clique para Executar)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-921">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="c3d2f-922">Histórico de atualizações do Office 2013 (clique para executar)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-922">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="c3d2f-923">Histórico de atualizações do Office 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-923">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="c3d2f-924">Histórico de atualizações do Outlook 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="c3d2f-924">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="c3d2f-925">Histórico de atualizações do Office para Mac</span><span class="sxs-lookup"><span data-stu-id="c3d2f-925">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="c3d2f-926">Criando Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="c3d2f-926">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)