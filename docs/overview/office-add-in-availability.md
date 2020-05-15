---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos com suporte para o Excel, OneNote, Outlook, PowerPoint, Project e Word.
ms.date: 05/11/2020
localization_priority: Priority
ms.openlocfilehash: 36c6bc6b6348ac988049f9a50127f6dd2f94bf37
ms.sourcegitcommit: 682d18c9149b1153f9c38d28e2a90384e6a261dc
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/13/2020
ms.locfileid: "44217820"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="b1beb-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="b1beb-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="b1beb-104">Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="b1beb-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="b1beb-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que são compatíveis atualmente com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="b1beb-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="b1beb-106">A versão inicial do Office 2016 instalada por meio do MSI apenas contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="b1beb-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="b1beb-107">Para saber mais sobre o histórico de atualizações de várias versões do Office, confira a seção[Confira também](#see-also).</span><span class="sxs-lookup"><span data-stu-id="b1beb-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="b1beb-108">Excel</span><span class="sxs-lookup"><span data-stu-id="b1beb-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="b1beb-109">Plataforma</span><span class="sxs-lookup"><span data-stu-id="b1beb-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="b1beb-110">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="b1beb-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="b1beb-111">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="b1beb-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="b1beb-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="b1beb-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-113">Office na Web</span><span class="sxs-lookup"><span data-stu-id="b1beb-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="b1beb-114">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b1beb-114">- TaskPane</span></span><br><span data-ttu-id="b1beb-115">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b1beb-115">
        - Content</span></span><br><span data-ttu-id="b1beb-116">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="b1beb-116">
        - Custom Functions</span></span><br><span data-ttu-id="b1beb-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="b1beb-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="b1beb-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b1beb-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b1beb-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b1beb-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b1beb-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b1beb-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b1beb-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b1beb-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b1beb-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="b1beb-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="b1beb-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="b1beb-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="b1beb-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="b1beb-131">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-131">
        - BindingEvents</span></span><br><span data-ttu-id="b1beb-132">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-132">
        - CompressedFile</span></span><br><span data-ttu-id="b1beb-133">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-133">
        - DocumentEvents</span></span><br><span data-ttu-id="b1beb-134">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b1beb-134">
        - File</span></span><br><span data-ttu-id="b1beb-135">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-135">
        - MatrixBindings</span></span><br><span data-ttu-id="b1beb-136">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-136">
        - MatrixCoercion</span></span><br><span data-ttu-id="b1beb-137">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="b1beb-137">
        - Selection</span></span><br><span data-ttu-id="b1beb-138">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="b1beb-138">
        - Settings</span></span><br><span data-ttu-id="b1beb-139">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-139">
        - TableBindings</span></span><br><span data-ttu-id="b1beb-140">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-140">
        - TableCoercion</span></span><br><span data-ttu-id="b1beb-141">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-141">
        - TextBindings</span></span><br><span data-ttu-id="b1beb-142">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-142">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-143">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="b1beb-143">Office on Windows</span></span><br><span data-ttu-id="b1beb-144">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b1beb-144">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b1beb-145">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b1beb-145">- TaskPane</span></span><br><span data-ttu-id="b1beb-146">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b1beb-146">
        - Content</span></span><br><span data-ttu-id="b1beb-147">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="b1beb-147">
        - Custom Functions</span></span><br><span data-ttu-id="b1beb-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="b1beb-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="b1beb-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b1beb-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b1beb-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b1beb-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b1beb-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b1beb-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b1beb-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b1beb-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b1beb-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="b1beb-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="b1beb-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="b1beb-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b1beb-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b1beb-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="b1beb-163">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-163">
        - BindingEvents</span></span><br><span data-ttu-id="b1beb-164">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-164">
        - CompressedFile</span></span><br><span data-ttu-id="b1beb-165">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-165">
        - DocumentEvents</span></span><br><span data-ttu-id="b1beb-166">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b1beb-166">
        - File</span></span><br><span data-ttu-id="b1beb-167">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-167">
        - MatrixBindings</span></span><br><span data-ttu-id="b1beb-168">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-168">
        - MatrixCoercion</span></span><br><span data-ttu-id="b1beb-169">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="b1beb-169">
        - Selection</span></span><br><span data-ttu-id="b1beb-170">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="b1beb-170">
        - Settings</span></span><br><span data-ttu-id="b1beb-171">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-171">
        - TableBindings</span></span><br><span data-ttu-id="b1beb-172">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-172">
        - TableCoercion</span></span><br><span data-ttu-id="b1beb-173">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-173">
        - TextBindings</span></span><br><span data-ttu-id="b1beb-174">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-174">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-175">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="b1beb-175">Office 2019 on Windows</span></span><br><span data-ttu-id="b1beb-176">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b1beb-176">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="b1beb-177">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b1beb-177">- TaskPane</span></span><br><span data-ttu-id="b1beb-178">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b1beb-178">
        - Content</span></span><br><span data-ttu-id="b1beb-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="b1beb-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b1beb-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b1beb-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b1beb-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b1beb-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b1beb-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b1beb-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b1beb-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b1beb-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b1beb-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b1beb-190">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-190">- BindingEvents</span></span><br><span data-ttu-id="b1beb-191">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-191">
        - CompressedFile</span></span><br><span data-ttu-id="b1beb-192">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-192">
        - DocumentEvents</span></span><br><span data-ttu-id="b1beb-193">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b1beb-193">
        - File</span></span><br><span data-ttu-id="b1beb-194">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-194">
        - MatrixBindings</span></span><br><span data-ttu-id="b1beb-195">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-195">
        - MatrixCoercion</span></span><br><span data-ttu-id="b1beb-196">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="b1beb-196">
        - Selection</span></span><br><span data-ttu-id="b1beb-197">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="b1beb-197">
        - Settings</span></span><br><span data-ttu-id="b1beb-198">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-198">
        - TableBindings</span></span><br><span data-ttu-id="b1beb-199">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-199">
        - TableCoercion</span></span><br><span data-ttu-id="b1beb-200">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-200">
        - TextBindings</span></span><br><span data-ttu-id="b1beb-201">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-201">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-202">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="b1beb-202">Office 2016 on Windows</span></span><br><span data-ttu-id="b1beb-203">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b1beb-203">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="b1beb-204">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b1beb-204">- TaskPane</span></span><br><span data-ttu-id="b1beb-205">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b1beb-205">
        - Content</span></span></td>
    <td><span data-ttu-id="b1beb-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b1beb-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="b1beb-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="b1beb-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b1beb-209">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-209">- BindingEvents</span></span><br><span data-ttu-id="b1beb-210">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-210">
        - CompressedFile</span></span><br><span data-ttu-id="b1beb-211">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-211">
        - DocumentEvents</span></span><br><span data-ttu-id="b1beb-212">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b1beb-212">
        - File</span></span><br><span data-ttu-id="b1beb-213">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-213">
        - MatrixBindings</span></span><br><span data-ttu-id="b1beb-214">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-214">
        - MatrixCoercion</span></span><br><span data-ttu-id="b1beb-215">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="b1beb-215">
        - Selection</span></span><br><span data-ttu-id="b1beb-216">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="b1beb-216">
        - Settings</span></span><br><span data-ttu-id="b1beb-217">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-217">
        - TableBindings</span></span><br><span data-ttu-id="b1beb-218">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-218">
        - TableCoercion</span></span><br><span data-ttu-id="b1beb-219">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-219">
        - TextBindings</span></span><br><span data-ttu-id="b1beb-220">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-220">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-221">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="b1beb-221">Office 2013 on Windows</span></span><br><span data-ttu-id="b1beb-222">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b1beb-222">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="b1beb-223">
        - TaskPane</span><span class="sxs-lookup"><span data-stu-id="b1beb-223">
        - TaskPane</span></span><br><span data-ttu-id="b1beb-224">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b1beb-224">
        - Content</span></span></td>
    <td>  <span data-ttu-id="b1beb-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b1beb-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="b1beb-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b1beb-227">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-227">
        - BindingEvents</span></span><br><span data-ttu-id="b1beb-228">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-228">
        - CompressedFile</span></span><br><span data-ttu-id="b1beb-229">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-229">
        - DocumentEvents</span></span><br><span data-ttu-id="b1beb-230">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b1beb-230">
        - File</span></span><br><span data-ttu-id="b1beb-231">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-231">
        - MatrixBindings</span></span><br><span data-ttu-id="b1beb-232">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-232">
        - MatrixCoercion</span></span><br><span data-ttu-id="b1beb-233">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="b1beb-233">
        - Selection</span></span><br><span data-ttu-id="b1beb-234">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="b1beb-234">
        - Settings</span></span><br><span data-ttu-id="b1beb-235">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-235">
        - TableBindings</span></span><br><span data-ttu-id="b1beb-236">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-236">
        - TableCoercion</span></span><br><span data-ttu-id="b1beb-237">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-237">
        - TextBindings</span></span><br><span data-ttu-id="b1beb-238">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-238">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-239">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="b1beb-239">Office on iPad</span></span><br><span data-ttu-id="b1beb-240">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b1beb-240">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="b1beb-241">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b1beb-241">- TaskPane</span></span><br><span data-ttu-id="b1beb-242">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b1beb-242">
        - Content</span></span></td>
    <td><span data-ttu-id="b1beb-243">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-243">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b1beb-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b1beb-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b1beb-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b1beb-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b1beb-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b1beb-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b1beb-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b1beb-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="b1beb-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="b1beb-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="b1beb-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b1beb-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b1beb-256">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-256">- BindingEvents</span></span><br><span data-ttu-id="b1beb-257">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-257">
        - DocumentEvents</span></span><br><span data-ttu-id="b1beb-258">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b1beb-258">
        - File</span></span><br><span data-ttu-id="b1beb-259">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-259">
        - MatrixBindings</span></span><br><span data-ttu-id="b1beb-260">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-260">
        - MatrixCoercion</span></span><br><span data-ttu-id="b1beb-261">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="b1beb-261">
        - Selection</span></span><br><span data-ttu-id="b1beb-262">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="b1beb-262">
        - Settings</span></span><br><span data-ttu-id="b1beb-263">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-263">
        - TableBindings</span></span><br><span data-ttu-id="b1beb-264">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-264">
        - TableCoercion</span></span><br><span data-ttu-id="b1beb-265">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-265">
        - TextBindings</span></span><br><span data-ttu-id="b1beb-266">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-266">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-267">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="b1beb-267">Office on Mac</span></span><br><span data-ttu-id="b1beb-268">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b1beb-268">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="b1beb-269">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b1beb-269">- TaskPane</span></span><br><span data-ttu-id="b1beb-270">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b1beb-270">
        - Content</span></span><br><span data-ttu-id="b1beb-271">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="b1beb-271">
        - Custom Functions</span></span><br><span data-ttu-id="b1beb-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="b1beb-273">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-273">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b1beb-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b1beb-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b1beb-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b1beb-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b1beb-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b1beb-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b1beb-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b1beb-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="b1beb-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="b1beb-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="b1beb-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b1beb-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b1beb-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="b1beb-287">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-287">- BindingEvents</span></span><br><span data-ttu-id="b1beb-288">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-288">
        - CompressedFile</span></span><br><span data-ttu-id="b1beb-289">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-289">
        - DocumentEvents</span></span><br><span data-ttu-id="b1beb-290">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b1beb-290">
        - File</span></span><br><span data-ttu-id="b1beb-291">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-291">
        - MatrixBindings</span></span><br><span data-ttu-id="b1beb-292">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-292">
        - MatrixCoercion</span></span><br><span data-ttu-id="b1beb-293">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-293">
        - PdfFile</span></span><br><span data-ttu-id="b1beb-294">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="b1beb-294">
        - Selection</span></span><br><span data-ttu-id="b1beb-295">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="b1beb-295">
        - Settings</span></span><br><span data-ttu-id="b1beb-296">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-296">
        - TableBindings</span></span><br><span data-ttu-id="b1beb-297">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-297">
        - TableCoercion</span></span><br><span data-ttu-id="b1beb-298">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-298">
        - TextBindings</span></span><br><span data-ttu-id="b1beb-299">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-299">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-300">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="b1beb-300">Office 2019 on Mac</span></span><br><span data-ttu-id="b1beb-301">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b1beb-301">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="b1beb-302">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b1beb-302">- TaskPane</span></span><br><span data-ttu-id="b1beb-303">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b1beb-303">
        - Content</span></span><br><span data-ttu-id="b1beb-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="b1beb-305">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-305">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b1beb-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b1beb-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b1beb-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b1beb-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b1beb-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b1beb-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b1beb-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b1beb-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b1beb-314">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-314">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b1beb-315">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-315">- BindingEvents</span></span><br><span data-ttu-id="b1beb-316">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-316">
        - CompressedFile</span></span><br><span data-ttu-id="b1beb-317">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-317">
        - DocumentEvents</span></span><br><span data-ttu-id="b1beb-318">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b1beb-318">
        - File</span></span><br><span data-ttu-id="b1beb-319">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-319">
        - MatrixBindings</span></span><br><span data-ttu-id="b1beb-320">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-320">
        - MatrixCoercion</span></span><br><span data-ttu-id="b1beb-321">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-321">
        - PdfFile</span></span><br><span data-ttu-id="b1beb-322">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="b1beb-322">
        - Selection</span></span><br><span data-ttu-id="b1beb-323">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="b1beb-323">
        - Settings</span></span><br><span data-ttu-id="b1beb-324">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-324">
        - TableBindings</span></span><br><span data-ttu-id="b1beb-325">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-325">
        - TableCoercion</span></span><br><span data-ttu-id="b1beb-326">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-326">
        - TextBindings</span></span><br><span data-ttu-id="b1beb-327">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-327">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-328">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="b1beb-328">Office 2016 on Mac</span></span><br><span data-ttu-id="b1beb-329">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b1beb-329">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="b1beb-330">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b1beb-330">- TaskPane</span></span><br><span data-ttu-id="b1beb-331">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b1beb-331">
        - Content</span></span></td>
    <td><span data-ttu-id="b1beb-332">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-332">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b1beb-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="b1beb-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="b1beb-334">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-334">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b1beb-335">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-335">- BindingEvents</span></span><br><span data-ttu-id="b1beb-336">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-336">
        - CompressedFile</span></span><br><span data-ttu-id="b1beb-337">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-337">
        - DocumentEvents</span></span><br><span data-ttu-id="b1beb-338">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b1beb-338">
        - File</span></span><br><span data-ttu-id="b1beb-339">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-339">
        - MatrixBindings</span></span><br><span data-ttu-id="b1beb-340">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-340">
        - MatrixCoercion</span></span><br><span data-ttu-id="b1beb-341">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-341">
        - PdfFile</span></span><br><span data-ttu-id="b1beb-342">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="b1beb-342">
        - Selection</span></span><br><span data-ttu-id="b1beb-343">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="b1beb-343">
        - Settings</span></span><br><span data-ttu-id="b1beb-344">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-344">
        - TableBindings</span></span><br><span data-ttu-id="b1beb-345">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-345">
        - TableCoercion</span></span><br><span data-ttu-id="b1beb-346">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-346">
        - TextBindings</span></span><br><span data-ttu-id="b1beb-347">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-347">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="b1beb-348">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="b1beb-348">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="b1beb-349">Funções personalizadas (somente Excel)</span><span class="sxs-lookup"><span data-stu-id="b1beb-349">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="b1beb-350">Plataforma</span><span class="sxs-lookup"><span data-stu-id="b1beb-350">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="b1beb-351">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="b1beb-351">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="b1beb-352">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="b1beb-352">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="b1beb-353"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="b1beb-353"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-354">Office na Web</span><span class="sxs-lookup"><span data-stu-id="b1beb-354">Office on the web</span></span></td>
    <td><span data-ttu-id="b1beb-355">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="b1beb-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="b1beb-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-357">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="b1beb-357">Office on Windows</span></span><br><span data-ttu-id="b1beb-358">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b1beb-358">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="b1beb-359">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="b1beb-359">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="b1beb-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-361">Office para Mac</span><span class="sxs-lookup"><span data-stu-id="b1beb-361">Office for Mac</span></span><br><span data-ttu-id="b1beb-362">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="b1beb-362">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="b1beb-363">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="b1beb-363">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="b1beb-364">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-364">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="b1beb-365">Outlook</span><span class="sxs-lookup"><span data-stu-id="b1beb-365">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b1beb-366">Plataforma</span><span class="sxs-lookup"><span data-stu-id="b1beb-366">Platform</span></span></th>
    <th><span data-ttu-id="b1beb-367">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="b1beb-367">Extension points</span></span></th>
    <th><span data-ttu-id="b1beb-368">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="b1beb-368">API requirement sets</span></span></th>
    <th><span data-ttu-id="b1beb-369"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="b1beb-369"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-370">Office na Web</span><span class="sxs-lookup"><span data-stu-id="b1beb-370">Office on the web</span></span><br><span data-ttu-id="b1beb-371">(moderno)</span><span class="sxs-lookup"><span data-stu-id="b1beb-371">(modern)</span></span></td>
    <td> <span data-ttu-id="b1beb-372">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-372">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="b1beb-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="b1beb-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="b1beb-375">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-375">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="b1beb-376">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-376">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b1beb-377">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-377">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b1beb-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b1beb-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b1beb-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b1beb-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b1beb-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b1beb-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="b1beb-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="b1beb-385">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b1beb-385">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-386">Office na Web</span><span class="sxs-lookup"><span data-stu-id="b1beb-386">Office on the web</span></span><br><span data-ttu-id="b1beb-387">(clássico)</span><span class="sxs-lookup"><span data-stu-id="b1beb-387">(classic)</span></span></td>
    <td> <span data-ttu-id="b1beb-388">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-388">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="b1beb-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="b1beb-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="b1beb-391">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-391">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="b1beb-392">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-392">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b1beb-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b1beb-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b1beb-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b1beb-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b1beb-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b1beb-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="b1beb-399">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b1beb-399">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-400">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="b1beb-400">Office on Windows</span></span><br><span data-ttu-id="b1beb-401">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b1beb-401">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b1beb-402">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-402">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="b1beb-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="b1beb-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="b1beb-405">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-405">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="b1beb-406">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-406">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="b1beb-407">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Módulos</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-407">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="b1beb-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b1beb-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b1beb-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b1beb-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b1beb-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b1beb-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b1beb-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="b1beb-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="b1beb-416">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b1beb-416">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-417">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="b1beb-417">Office 2019 on Windows</span></span><br><span data-ttu-id="b1beb-418">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b1beb-418">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b1beb-419">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-419">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="b1beb-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="b1beb-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="b1beb-422">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-422">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="b1beb-423">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-423">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="b1beb-424">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Módulos</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-424">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="b1beb-425">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-425">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b1beb-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b1beb-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b1beb-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b1beb-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b1beb-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b1beb-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="b1beb-432">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b1beb-432">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-433">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="b1beb-433">Office 2016 on Windows</span></span><br><span data-ttu-id="b1beb-434">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b1beb-434">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b1beb-435">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-435">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="b1beb-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="b1beb-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="b1beb-438">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-438">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="b1beb-439">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-439">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="b1beb-440">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Módulos</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-440">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="b1beb-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b1beb-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b1beb-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b1beb-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="b1beb-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="b1beb-445">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b1beb-445">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-446">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="b1beb-446">Office 2013 on Windows</span></span><br><span data-ttu-id="b1beb-447">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b1beb-447">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b1beb-448">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-448">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="b1beb-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="b1beb-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="b1beb-451">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-451">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br>
    <td> <span data-ttu-id="b1beb-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b1beb-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b1beb-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="b1beb-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="b1beb-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="b1beb-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="b1beb-456">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b1beb-456">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-457">Office no iOS</span><span class="sxs-lookup"><span data-stu-id="b1beb-457">Office on iOS</span></span><br><span data-ttu-id="b1beb-458">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b1beb-458">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b1beb-459">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-459">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="b1beb-460">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-460">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b1beb-461">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-461">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b1beb-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b1beb-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b1beb-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b1beb-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="b1beb-466">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b1beb-466">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-467">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="b1beb-467">Office on Mac</span></span><br><span data-ttu-id="b1beb-468">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b1beb-468">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b1beb-469">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-469">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="b1beb-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="b1beb-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="b1beb-472">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-472">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="b1beb-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b1beb-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b1beb-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b1beb-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b1beb-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b1beb-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b1beb-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b1beb-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="b1beb-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="b1beb-482">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b1beb-482">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-483">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="b1beb-483">Office 2019 on Mac</span></span><br><span data-ttu-id="b1beb-484">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b1beb-484">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b1beb-485">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-485">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="b1beb-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="b1beb-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="b1beb-488">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-488">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="b1beb-489">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-489">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b1beb-490">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-490">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b1beb-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b1beb-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b1beb-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b1beb-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b1beb-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="b1beb-496">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b1beb-496">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-497">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="b1beb-497">Office 2016 on Mac</span></span><br><span data-ttu-id="b1beb-498">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b1beb-498">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b1beb-499">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-499">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="b1beb-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="b1beb-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="b1beb-502">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-502">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="b1beb-503">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-503">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b1beb-504">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-504">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b1beb-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b1beb-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b1beb-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b1beb-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b1beb-509">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-509">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="b1beb-510">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b1beb-510">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-511">Outlook no Android</span><span class="sxs-lookup"><span data-stu-id="b1beb-511">Office on Android</span></span><br><span data-ttu-id="b1beb-512">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b1beb-512">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b1beb-513">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-513">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="b1beb-514">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Organizador de compromissos (Redigir): reunião on-line (visualização)</span><span class="sxs-lookup"><span data-stu-id="b1beb-514">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Appointment Organizer (Compose): online meeting</a> (preview)</span></span><br><span data-ttu-id="b1beb-515">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-515">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b1beb-516">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-516">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b1beb-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b1beb-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b1beb-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b1beb-520">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-520">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="b1beb-521">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b1beb-521">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="b1beb-522">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="b1beb-522">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b1beb-523">O suporte ao cliente para um conjunto de requisitos pode ser restringido pelo suporte do servidor Exchange.</span><span class="sxs-lookup"><span data-stu-id="b1beb-523">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="b1beb-524">Consulte [Conjuntos de requisitos da API JavaScript do Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) para obter detalhes sobre o intervalo de conjuntos de requisitos suportado pelo servidor Exchange e pelos clientes Outlook.</span><span class="sxs-lookup"><span data-stu-id="b1beb-524">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="b1beb-525">Word</span><span class="sxs-lookup"><span data-stu-id="b1beb-525">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b1beb-526">Plataforma</span><span class="sxs-lookup"><span data-stu-id="b1beb-526">Platform</span></span></th>
    <th><span data-ttu-id="b1beb-527">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="b1beb-527">Extension points</span></span></th>
    <th><span data-ttu-id="b1beb-528">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="b1beb-528">API requirement sets</span></span></th>
    <th><span data-ttu-id="b1beb-529"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="b1beb-529"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-530">Office na Web</span><span class="sxs-lookup"><span data-stu-id="b1beb-530">Office on the web</span></span></td>
    <td> <span data-ttu-id="b1beb-531">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b1beb-531">- TaskPane</span></span><br><span data-ttu-id="b1beb-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b1beb-533">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-533">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b1beb-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="b1beb-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="b1beb-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b1beb-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b1beb-538">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-538">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="b1beb-539">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-539">- BindingEvents</span></span><br><span data-ttu-id="b1beb-540">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b1beb-540">
         - CustomXmlParts</span></span><br><span data-ttu-id="b1beb-541">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-541">
         - DocumentEvents</span></span><br><span data-ttu-id="b1beb-542">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b1beb-542">
         - File</span></span><br><span data-ttu-id="b1beb-543">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-543">
         - HtmlCoercion</span></span><br><span data-ttu-id="b1beb-544">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-544">
         - MatrixBindings</span></span><br><span data-ttu-id="b1beb-545">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-545">
         - MatrixCoercion</span></span><br><span data-ttu-id="b1beb-546">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-546">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b1beb-547">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-547">
         - PdfFile</span></span><br><span data-ttu-id="b1beb-548">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b1beb-548">
         - Selection</span></span><br><span data-ttu-id="b1beb-549">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b1beb-549">
         - Settings</span></span><br><span data-ttu-id="b1beb-550">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-550">
         - TableBindings</span></span><br><span data-ttu-id="b1beb-551">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-551">
         - TableCoercion</span></span><br><span data-ttu-id="b1beb-552">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-552">
         - TextBindings</span></span><br><span data-ttu-id="b1beb-553">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-553">
         - TextCoercion</span></span><br><span data-ttu-id="b1beb-554">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-554">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-555">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="b1beb-555">Office on Windows</span></span><br><span data-ttu-id="b1beb-556">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b1beb-556">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b1beb-557">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b1beb-557">- TaskPane</span></span><br><span data-ttu-id="b1beb-558">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-558">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b1beb-559">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-559">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b1beb-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="b1beb-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="b1beb-562">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-562">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b1beb-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b1beb-564">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-564">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="b1beb-565">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-565">- BindingEvents</span></span><br><span data-ttu-id="b1beb-566">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-566">
         - CompressedFile</span></span><br><span data-ttu-id="b1beb-567">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b1beb-567">
         - CustomXmlParts</span></span><br><span data-ttu-id="b1beb-568">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-568">
         - DocumentEvents</span></span><br><span data-ttu-id="b1beb-569">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b1beb-569">
         - File</span></span><br><span data-ttu-id="b1beb-570">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-570">
         - HtmlCoercion</span></span><br><span data-ttu-id="b1beb-571">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-571">
         - MatrixBindings</span></span><br><span data-ttu-id="b1beb-572">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-572">
         - MatrixCoercion</span></span><br><span data-ttu-id="b1beb-573">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-573">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b1beb-574">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-574">
         - PdfFile</span></span><br><span data-ttu-id="b1beb-575">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b1beb-575">
         - Selection</span></span><br><span data-ttu-id="b1beb-576">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b1beb-576">
         - Settings</span></span><br><span data-ttu-id="b1beb-577">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-577">
         - TableBindings</span></span><br><span data-ttu-id="b1beb-578">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-578">
         - TableCoercion</span></span><br><span data-ttu-id="b1beb-579">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-579">
         - TextBindings</span></span><br><span data-ttu-id="b1beb-580">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-580">
         - TextCoercion</span></span><br><span data-ttu-id="b1beb-581">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-581">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-582">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="b1beb-582">Office 2019 on Windows</span></span><br><span data-ttu-id="b1beb-583">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b1beb-583">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b1beb-584">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b1beb-584">- TaskPane</span></span><br><span data-ttu-id="b1beb-585">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-585">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b1beb-586">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-586">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b1beb-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="b1beb-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="b1beb-589">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-589">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b1beb-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b1beb-591">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-591">- BindingEvents</span></span><br><span data-ttu-id="b1beb-592">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-592">
         - CompressedFile</span></span><br><span data-ttu-id="b1beb-593">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b1beb-593">
         - CustomXmlParts</span></span><br><span data-ttu-id="b1beb-594">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-594">
         - DocumentEvents</span></span><br><span data-ttu-id="b1beb-595">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b1beb-595">
         - File</span></span><br><span data-ttu-id="b1beb-596">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-596">
         - HtmlCoercion</span></span><br><span data-ttu-id="b1beb-597">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-597">
         - MatrixBindings</span></span><br><span data-ttu-id="b1beb-598">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-598">
         - MatrixCoercion</span></span><br><span data-ttu-id="b1beb-599">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-599">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b1beb-600">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-600">
         - PdfFile</span></span><br><span data-ttu-id="b1beb-601">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b1beb-601">
         - Selection</span></span><br><span data-ttu-id="b1beb-602">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b1beb-602">
         - Settings</span></span><br><span data-ttu-id="b1beb-603">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-603">
         - TableBindings</span></span><br><span data-ttu-id="b1beb-604">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-604">
         - TableCoercion</span></span><br><span data-ttu-id="b1beb-605">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-605">
         - TextBindings</span></span><br><span data-ttu-id="b1beb-606">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-606">
         - TextCoercion</span></span><br><span data-ttu-id="b1beb-607">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-607">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-608">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="b1beb-608">Office 2016 on Windows</span></span><br><span data-ttu-id="b1beb-609">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b1beb-609">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b1beb-610">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b1beb-610">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b1beb-611">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-611">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b1beb-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="b1beb-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="b1beb-613">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-613">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b1beb-614">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-614">- BindingEvents</span></span><br><span data-ttu-id="b1beb-615">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-615">
         - CompressedFile</span></span><br><span data-ttu-id="b1beb-616">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b1beb-616">
         - CustomXmlParts</span></span><br><span data-ttu-id="b1beb-617">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-617">
         - DocumentEvents</span></span><br><span data-ttu-id="b1beb-618">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b1beb-618">
         - File</span></span><br><span data-ttu-id="b1beb-619">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-619">
         - HtmlCoercion</span></span><br><span data-ttu-id="b1beb-620">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-620">
         - MatrixBindings</span></span><br><span data-ttu-id="b1beb-621">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-621">
         - MatrixCoercion</span></span><br><span data-ttu-id="b1beb-622">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-622">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b1beb-623">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-623">
         - PdfFile</span></span><br><span data-ttu-id="b1beb-624">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b1beb-624">
         - Selection</span></span><br><span data-ttu-id="b1beb-625">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b1beb-625">
         - Settings</span></span><br><span data-ttu-id="b1beb-626">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-626">
         - TableBindings</span></span><br><span data-ttu-id="b1beb-627">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-627">
         - TableCoercion</span></span><br><span data-ttu-id="b1beb-628">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-628">
         - TextBindings</span></span><br><span data-ttu-id="b1beb-629">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-629">
         - TextCoercion</span></span><br><span data-ttu-id="b1beb-630">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-630">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-631">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="b1beb-631">Office 2013 on Windows</span></span><br><span data-ttu-id="b1beb-632">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b1beb-632">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b1beb-633">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b1beb-633">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b1beb-634">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b1beb-634">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="b1beb-635">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-635">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b1beb-636">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-636">- BindingEvents</span></span><br><span data-ttu-id="b1beb-637">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-637">
         - CompressedFile</span></span><br><span data-ttu-id="b1beb-638">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b1beb-638">
         - CustomXmlParts</span></span><br><span data-ttu-id="b1beb-639">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-639">
         - DocumentEvents</span></span><br><span data-ttu-id="b1beb-640">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b1beb-640">
         - File</span></span><br><span data-ttu-id="b1beb-641">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-641">
         - HtmlCoercion</span></span><br><span data-ttu-id="b1beb-642">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-642">
         - MatrixBindings</span></span><br><span data-ttu-id="b1beb-643">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-643">
         - MatrixCoercion</span></span><br><span data-ttu-id="b1beb-644">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-644">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b1beb-645">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-645">
         - PdfFile</span></span><br><span data-ttu-id="b1beb-646">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b1beb-646">
         - Selection</span></span><br><span data-ttu-id="b1beb-647">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b1beb-647">
         - Settings</span></span><br><span data-ttu-id="b1beb-648">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-648">
         - TableBindings</span></span><br><span data-ttu-id="b1beb-649">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-649">
         - TableCoercion</span></span><br><span data-ttu-id="b1beb-650">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-650">
         - TextBindings</span></span><br><span data-ttu-id="b1beb-651">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-651">
         - TextCoercion</span></span><br><span data-ttu-id="b1beb-652">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-652">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-653">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="b1beb-653">Office on iPad</span></span><br><span data-ttu-id="b1beb-654">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b1beb-654">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b1beb-655">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b1beb-655">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b1beb-656">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-656">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b1beb-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="b1beb-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="b1beb-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b1beb-660">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-660">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="b1beb-661">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-661">- BindingEvents</span></span><br><span data-ttu-id="b1beb-662">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-662">
         - CompressedFile</span></span><br><span data-ttu-id="b1beb-663">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b1beb-663">
         - CustomXmlParts</span></span><br><span data-ttu-id="b1beb-664">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-664">
         - DocumentEvents</span></span><br><span data-ttu-id="b1beb-665">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b1beb-665">
         - File</span></span><br><span data-ttu-id="b1beb-666">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-666">
         - HtmlCoercion</span></span><br><span data-ttu-id="b1beb-667">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-667">
         - MatrixBindings</span></span><br><span data-ttu-id="b1beb-668">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-668">
         - MatrixCoercion</span></span><br><span data-ttu-id="b1beb-669">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-669">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b1beb-670">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-670">
         - PdfFile</span></span><br><span data-ttu-id="b1beb-671">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b1beb-671">
         - Selection</span></span><br><span data-ttu-id="b1beb-672">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b1beb-672">
         - Settings</span></span><br><span data-ttu-id="b1beb-673">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-673">
         - TableBindings</span></span><br><span data-ttu-id="b1beb-674">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-674">
         - TableCoercion</span></span><br><span data-ttu-id="b1beb-675">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-675">
         - TextBindings</span></span><br><span data-ttu-id="b1beb-676">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-676">
         - TextCoercion</span></span><br><span data-ttu-id="b1beb-677">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-677">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-678">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="b1beb-678">Office on Mac</span></span><br><span data-ttu-id="b1beb-679">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b1beb-679">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b1beb-680">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b1beb-680">- TaskPane</span></span><br><span data-ttu-id="b1beb-681">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-681">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b1beb-682">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-682">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b1beb-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="b1beb-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="b1beb-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b1beb-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b1beb-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="b1beb-688">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-688">- BindingEvents</span></span><br><span data-ttu-id="b1beb-689">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-689">
         - CompressedFile</span></span><br><span data-ttu-id="b1beb-690">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b1beb-690">
         - CustomXmlParts</span></span><br><span data-ttu-id="b1beb-691">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-691">
         - DocumentEvents</span></span><br><span data-ttu-id="b1beb-692">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b1beb-692">
         - File</span></span><br><span data-ttu-id="b1beb-693">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-693">
         - HtmlCoercion</span></span><br><span data-ttu-id="b1beb-694">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-694">
         - MatrixBindings</span></span><br><span data-ttu-id="b1beb-695">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-695">
         - MatrixCoercion</span></span><br><span data-ttu-id="b1beb-696">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-696">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b1beb-697">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-697">
         - PdfFile</span></span><br><span data-ttu-id="b1beb-698">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b1beb-698">
         - Selection</span></span><br><span data-ttu-id="b1beb-699">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b1beb-699">
         - Settings</span></span><br><span data-ttu-id="b1beb-700">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-700">
         - TableBindings</span></span><br><span data-ttu-id="b1beb-701">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-701">
         - TableCoercion</span></span><br><span data-ttu-id="b1beb-702">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-702">
         - TextBindings</span></span><br><span data-ttu-id="b1beb-703">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-703">
         - TextCoercion</span></span><br><span data-ttu-id="b1beb-704">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-704">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-705">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="b1beb-705">Office 2019 on Mac</span></span><br><span data-ttu-id="b1beb-706">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b1beb-706">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b1beb-707">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b1beb-707">- TaskPane</span></span><br><span data-ttu-id="b1beb-708">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-708">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b1beb-709">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-709">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b1beb-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="b1beb-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="b1beb-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b1beb-713">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-713">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="b1beb-714">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-714">- BindingEvents</span></span><br><span data-ttu-id="b1beb-715">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-715">
         - CompressedFile</span></span><br><span data-ttu-id="b1beb-716">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b1beb-716">
         - CustomXmlParts</span></span><br><span data-ttu-id="b1beb-717">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-717">
         - DocumentEvents</span></span><br><span data-ttu-id="b1beb-718">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b1beb-718">
         - File</span></span><br><span data-ttu-id="b1beb-719">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-719">
         - HtmlCoercion</span></span><br><span data-ttu-id="b1beb-720">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-720">
         - MatrixBindings</span></span><br><span data-ttu-id="b1beb-721">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-721">
         - MatrixCoercion</span></span><br><span data-ttu-id="b1beb-722">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-722">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b1beb-723">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-723">
         - PdfFile</span></span><br><span data-ttu-id="b1beb-724">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b1beb-724">
         - Selection</span></span><br><span data-ttu-id="b1beb-725">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b1beb-725">
         - Settings</span></span><br><span data-ttu-id="b1beb-726">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-726">
         - TableBindings</span></span><br><span data-ttu-id="b1beb-727">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-727">
         - TableCoercion</span></span><br><span data-ttu-id="b1beb-728">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-728">
         - TextBindings</span></span><br><span data-ttu-id="b1beb-729">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-729">
         - TextCoercion</span></span><br><span data-ttu-id="b1beb-730">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-730">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-731">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="b1beb-731">Office 2016 on Mac</span></span><br><span data-ttu-id="b1beb-732">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b1beb-732">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b1beb-733">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b1beb-733">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b1beb-734">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-734">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b1beb-735">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="b1beb-735">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="b1beb-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b1beb-737">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-737">- BindingEvents</span></span><br><span data-ttu-id="b1beb-738">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-738">
         - CompressedFile</span></span><br><span data-ttu-id="b1beb-739">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b1beb-739">
         - CustomXmlParts</span></span><br><span data-ttu-id="b1beb-740">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-740">
         - DocumentEvents</span></span><br><span data-ttu-id="b1beb-741">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b1beb-741">
         - File</span></span><br><span data-ttu-id="b1beb-742">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-742">
         - HtmlCoercion</span></span><br><span data-ttu-id="b1beb-743">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-743">
         - MatrixBindings</span></span><br><span data-ttu-id="b1beb-744">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-744">
         - MatrixCoercion</span></span><br><span data-ttu-id="b1beb-745">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-745">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b1beb-746">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-746">
         - PdfFile</span></span><br><span data-ttu-id="b1beb-747">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b1beb-747">
         - Selection</span></span><br><span data-ttu-id="b1beb-748">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b1beb-748">
         - Settings</span></span><br><span data-ttu-id="b1beb-749">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-749">
         - TableBindings</span></span><br><span data-ttu-id="b1beb-750">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-750">
         - TableCoercion</span></span><br><span data-ttu-id="b1beb-751">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b1beb-751">
         - TextBindings</span></span><br><span data-ttu-id="b1beb-752">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-752">
         - TextCoercion</span></span><br><span data-ttu-id="b1beb-753">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-753">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="b1beb-754">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="b1beb-754">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="b1beb-755">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="b1beb-755">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b1beb-756">Plataforma</span><span class="sxs-lookup"><span data-stu-id="b1beb-756">Platform</span></span></th>
    <th><span data-ttu-id="b1beb-757">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="b1beb-757">Extension points</span></span></th>
    <th><span data-ttu-id="b1beb-758">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="b1beb-758">API requirement sets</span></span></th>
    <th><span data-ttu-id="b1beb-759"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="b1beb-759"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-760">Office na Web</span><span class="sxs-lookup"><span data-stu-id="b1beb-760">Office on the web</span></span></td>
    <td> <span data-ttu-id="b1beb-761">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b1beb-761">- Content</span></span><br><span data-ttu-id="b1beb-762">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="b1beb-762">
         - TaskPane</span></span><br><span data-ttu-id="b1beb-763">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-763">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b1beb-764">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-764">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="b1beb-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b1beb-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b1beb-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="b1beb-768">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b1beb-768">- ActiveView</span></span><br><span data-ttu-id="b1beb-769">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-769">
         - CompressedFile</span></span><br><span data-ttu-id="b1beb-770">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-770">
         - DocumentEvents</span></span><br><span data-ttu-id="b1beb-771">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b1beb-771">
         - File</span></span><br><span data-ttu-id="b1beb-772">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-772">
         - PdfFile</span></span><br><span data-ttu-id="b1beb-773">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b1beb-773">
         - Selection</span></span><br><span data-ttu-id="b1beb-774">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b1beb-774">
         - Settings</span></span><br><span data-ttu-id="b1beb-775">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-775">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-776">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="b1beb-776">Office on Windows</span></span><br><span data-ttu-id="b1beb-777">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b1beb-777">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b1beb-778">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b1beb-778">- Content</span></span><br><span data-ttu-id="b1beb-779">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="b1beb-779">
         - TaskPane</span></span><br><span data-ttu-id="b1beb-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b1beb-781">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-781">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="b1beb-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b1beb-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b1beb-784">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-784">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="b1beb-785">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b1beb-785">- ActiveView</span></span><br><span data-ttu-id="b1beb-786">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-786">
         - CompressedFile</span></span><br><span data-ttu-id="b1beb-787">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-787">
         - DocumentEvents</span></span><br><span data-ttu-id="b1beb-788">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b1beb-788">
         - File</span></span><br><span data-ttu-id="b1beb-789">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-789">
         - PdfFile</span></span><br><span data-ttu-id="b1beb-790">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b1beb-790">
         - Selection</span></span><br><span data-ttu-id="b1beb-791">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b1beb-791">
         - Settings</span></span><br><span data-ttu-id="b1beb-792">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-792">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-793">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="b1beb-793">Office 2019 on Windows</span></span><br><span data-ttu-id="b1beb-794">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b1beb-794">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b1beb-795">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b1beb-795">- Content</span></span><br><span data-ttu-id="b1beb-796">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="b1beb-796">
         - TaskPane</span></span><br><span data-ttu-id="b1beb-797">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-797">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b1beb-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b1beb-799">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-799">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b1beb-800">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b1beb-800">- ActiveView</span></span><br><span data-ttu-id="b1beb-801">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-801">
         - CompressedFile</span></span><br><span data-ttu-id="b1beb-802">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-802">
         - DocumentEvents</span></span><br><span data-ttu-id="b1beb-803">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b1beb-803">
         - File</span></span><br><span data-ttu-id="b1beb-804">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-804">
         - PdfFile</span></span><br><span data-ttu-id="b1beb-805">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b1beb-805">
         - Selection</span></span><br><span data-ttu-id="b1beb-806">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b1beb-806">
         - Settings</span></span><br><span data-ttu-id="b1beb-807">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-807">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-808">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="b1beb-808">Office 2016 on Windows</span></span><br><span data-ttu-id="b1beb-809">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b1beb-809">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b1beb-810">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b1beb-810">- Content</span></span><br><span data-ttu-id="b1beb-811">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="b1beb-811">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="b1beb-812">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b1beb-812">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="b1beb-813">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-813">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b1beb-814">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b1beb-814">- ActiveView</span></span><br><span data-ttu-id="b1beb-815">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-815">
         - CompressedFile</span></span><br><span data-ttu-id="b1beb-816">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-816">
         - DocumentEvents</span></span><br><span data-ttu-id="b1beb-817">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b1beb-817">
         - File</span></span><br><span data-ttu-id="b1beb-818">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-818">
         - PdfFile</span></span><br><span data-ttu-id="b1beb-819">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b1beb-819">
         - Selection</span></span><br><span data-ttu-id="b1beb-820">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b1beb-820">
         - Settings</span></span><br><span data-ttu-id="b1beb-821">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-821">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-822">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="b1beb-822">Office 2013 on Windows</span></span><br><span data-ttu-id="b1beb-823">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b1beb-823">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b1beb-824">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b1beb-824">- Content</span></span><br><span data-ttu-id="b1beb-825">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="b1beb-825">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="b1beb-826">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b1beb-826">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="b1beb-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b1beb-828">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b1beb-828">- ActiveView</span></span><br><span data-ttu-id="b1beb-829">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-829">
         - CompressedFile</span></span><br><span data-ttu-id="b1beb-830">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-830">
         - DocumentEvents</span></span><br><span data-ttu-id="b1beb-831">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b1beb-831">
         - File</span></span><br><span data-ttu-id="b1beb-832">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-832">
         - PdfFile</span></span><br><span data-ttu-id="b1beb-833">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b1beb-833">
         - Selection</span></span><br><span data-ttu-id="b1beb-834">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b1beb-834">
         - Settings</span></span><br><span data-ttu-id="b1beb-835">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-835">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-836">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="b1beb-836">Office on iPad</span></span><br><span data-ttu-id="b1beb-837">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b1beb-837">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b1beb-838">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b1beb-838">- Content</span></span><br><span data-ttu-id="b1beb-839">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="b1beb-839">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="b1beb-840">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-840">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="b1beb-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b1beb-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b1beb-843">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b1beb-843">- ActiveView</span></span><br><span data-ttu-id="b1beb-844">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-844">
         - CompressedFile</span></span><br><span data-ttu-id="b1beb-845">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-845">
         - DocumentEvents</span></span><br><span data-ttu-id="b1beb-846">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b1beb-846">
         - File</span></span><br><span data-ttu-id="b1beb-847">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-847">
         - PdfFile</span></span><br><span data-ttu-id="b1beb-848">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b1beb-848">
         - Selection</span></span><br><span data-ttu-id="b1beb-849">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b1beb-849">
         - Settings</span></span><br><span data-ttu-id="b1beb-850">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-850">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-851">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="b1beb-851">Office on Mac</span></span><br><span data-ttu-id="b1beb-852">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b1beb-852">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b1beb-853">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b1beb-853">- Content</span></span><br><span data-ttu-id="b1beb-854">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="b1beb-854">
         - TaskPane</span></span><br><span data-ttu-id="b1beb-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b1beb-856">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-856">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="b1beb-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b1beb-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b1beb-859">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-859">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="b1beb-860">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b1beb-860">- ActiveView</span></span><br><span data-ttu-id="b1beb-861">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-861">
         - CompressedFile</span></span><br><span data-ttu-id="b1beb-862">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-862">
         - DocumentEvents</span></span><br><span data-ttu-id="b1beb-863">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b1beb-863">
         - File</span></span><br><span data-ttu-id="b1beb-864">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-864">
         - PdfFile</span></span><br><span data-ttu-id="b1beb-865">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b1beb-865">
         - Selection</span></span><br><span data-ttu-id="b1beb-866">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b1beb-866">
         - Settings</span></span><br><span data-ttu-id="b1beb-867">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-867">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-868">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="b1beb-868">Office 2019 on Mac</span></span><br><span data-ttu-id="b1beb-869">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b1beb-869">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b1beb-870">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b1beb-870">- Content</span></span><br><span data-ttu-id="b1beb-871">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="b1beb-871">
         - TaskPane</span></span><br><span data-ttu-id="b1beb-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b1beb-873">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-873">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b1beb-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b1beb-875">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b1beb-875">- ActiveView</span></span><br><span data-ttu-id="b1beb-876">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-876">
         - CompressedFile</span></span><br><span data-ttu-id="b1beb-877">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-877">
         - DocumentEvents</span></span><br><span data-ttu-id="b1beb-878">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b1beb-878">
         - File</span></span><br><span data-ttu-id="b1beb-879">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-879">
         - PdfFile</span></span><br><span data-ttu-id="b1beb-880">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b1beb-880">
         - Selection</span></span><br><span data-ttu-id="b1beb-881">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b1beb-881">
         - Settings</span></span><br><span data-ttu-id="b1beb-882">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-882">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-883">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="b1beb-883">Office 2016 on Mac</span></span><br><span data-ttu-id="b1beb-884">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b1beb-884">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b1beb-885">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b1beb-885">- Content</span></span><br><span data-ttu-id="b1beb-886">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="b1beb-886">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="b1beb-887">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b1beb-887">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="b1beb-888">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-888">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b1beb-889">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b1beb-889">- ActiveView</span></span><br><span data-ttu-id="b1beb-890">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-890">
         - CompressedFile</span></span><br><span data-ttu-id="b1beb-891">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-891">
         - DocumentEvents</span></span><br><span data-ttu-id="b1beb-892">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b1beb-892">
         - File</span></span><br><span data-ttu-id="b1beb-893">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1beb-893">
         - PdfFile</span></span><br><span data-ttu-id="b1beb-894">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b1beb-894">
         - Selection</span></span><br><span data-ttu-id="b1beb-895">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b1beb-895">
         - Settings</span></span><br><span data-ttu-id="b1beb-896">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-896">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="b1beb-897">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="b1beb-897">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="b1beb-898">OneNote</span><span class="sxs-lookup"><span data-stu-id="b1beb-898">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b1beb-899">Plataforma</span><span class="sxs-lookup"><span data-stu-id="b1beb-899">Platform</span></span></th>
    <th><span data-ttu-id="b1beb-900">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="b1beb-900">Extension points</span></span></th>
    <th><span data-ttu-id="b1beb-901">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="b1beb-901">API requirement sets</span></span></th>
    <th><span data-ttu-id="b1beb-902"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="b1beb-902"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-903">Office na Web</span><span class="sxs-lookup"><span data-stu-id="b1beb-903">Office on the web</span></span></td>
    <td> <span data-ttu-id="b1beb-904">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b1beb-904">- Content</span></span><br><span data-ttu-id="b1beb-905">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="b1beb-905">
         - TaskPane</span></span><br><span data-ttu-id="b1beb-906">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-906">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b1beb-907">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-907">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="b1beb-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b1beb-909">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-909">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b1beb-910">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1beb-910">- DocumentEvents</span></span><br><span data-ttu-id="b1beb-911">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-911">
         - HtmlCoercion</span></span><br><span data-ttu-id="b1beb-912">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b1beb-912">
         - Settings</span></span><br><span data-ttu-id="b1beb-913">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-913">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="b1beb-914">Project</span><span class="sxs-lookup"><span data-stu-id="b1beb-914">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b1beb-915">Plataforma</span><span class="sxs-lookup"><span data-stu-id="b1beb-915">Platform</span></span></th>
    <th><span data-ttu-id="b1beb-916">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="b1beb-916">Extension points</span></span></th>
    <th><span data-ttu-id="b1beb-917">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="b1beb-917">API requirement sets</span></span></th>
    <th><span data-ttu-id="b1beb-918"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="b1beb-918"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-919">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="b1beb-919">Office 2019 on Windows</span></span><br><span data-ttu-id="b1beb-920">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b1beb-920">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b1beb-921">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b1beb-921">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b1beb-922">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-922">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b1beb-923">- Seleção</span><span class="sxs-lookup"><span data-stu-id="b1beb-923">- Selection</span></span><br><span data-ttu-id="b1beb-924">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-924">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-925">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="b1beb-925">Office 2016 on Windows</span></span><br><span data-ttu-id="b1beb-926">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b1beb-926">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b1beb-927">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b1beb-927">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b1beb-928">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-928">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b1beb-929">- Seleção</span><span class="sxs-lookup"><span data-stu-id="b1beb-929">- Selection</span></span><br><span data-ttu-id="b1beb-930">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-930">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1beb-931">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="b1beb-931">Office 2013 on Windows</span></span><br><span data-ttu-id="b1beb-932">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b1beb-932">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b1beb-933">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b1beb-933">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b1beb-934">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1beb-934">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b1beb-935">- Seleção</span><span class="sxs-lookup"><span data-stu-id="b1beb-935">- Selection</span></span><br><span data-ttu-id="b1beb-936">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1beb-936">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="b1beb-937">Confira também</span><span class="sxs-lookup"><span data-stu-id="b1beb-937">See also</span></span>

- [<span data-ttu-id="b1beb-938">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="b1beb-938">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="b1beb-939">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="b1beb-939">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="b1beb-940">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="b1beb-940">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="b1beb-941">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="b1beb-941">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="b1beb-942">Documentação de Referência da API</span><span class="sxs-lookup"><span data-stu-id="b1beb-942">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="b1beb-943">Histórico de atualizações do Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="b1beb-943">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="b1beb-944">Histórico de atualizações do Office 2016 e 2019 (Clique para Executar)</span><span class="sxs-lookup"><span data-stu-id="b1beb-944">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="b1beb-945">Histórico de atualizações do Office 2013 (clique para executar)</span><span class="sxs-lookup"><span data-stu-id="b1beb-945">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="b1beb-946">Histórico de atualizações do Office 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="b1beb-946">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="b1beb-947">Histórico de atualizações do Outlook 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="b1beb-947">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="b1beb-948">Histórico de atualizações do Office para Mac</span><span class="sxs-lookup"><span data-stu-id="b1beb-948">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="b1beb-949">Criando Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="b1beb-949">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)