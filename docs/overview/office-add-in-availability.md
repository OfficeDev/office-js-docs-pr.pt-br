---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos com suporte para o Excel, OneNote, Outlook, PowerPoint, Project e Word.
ms.date: 10/09/2019
localization_priority: Priority
ms.openlocfilehash: 28d63866a03bcae99829d3a6b6c6198059a92bdc
ms.sourcegitcommit: 4d9f3e177b0bcd62804d5045f52b03e441af244f
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/10/2019
ms.locfileid: "37440147"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="1b6e9-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="1b6e9-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="1b6e9-104">Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="1b6e9-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="1b6e9-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que são compatíveis atualmente com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="1b6e9-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="1b6e9-106">A versão inicial do Office 2016 instalada por meio do MSI apenas contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="1b6e9-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="1b6e9-107">Para saber mais sobre o histórico de atualizações de várias versões do Office, confira a seção[Confira também](#see-also).</span><span class="sxs-lookup"><span data-stu-id="1b6e9-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="1b6e9-108">Excel</span><span class="sxs-lookup"><span data-stu-id="1b6e9-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="1b6e9-109">Plataforma</span><span class="sxs-lookup"><span data-stu-id="1b6e9-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="1b6e9-110">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="1b6e9-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="1b6e9-111">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="1b6e9-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="1b6e9-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-113">Office na Web</span><span class="sxs-lookup"><span data-stu-id="1b6e9-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="1b6e9-114">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="1b6e9-114">- TaskPane</span></span><br><span data-ttu-id="1b6e9-115">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-115">
        - Content</span></span><br><span data-ttu-id="1b6e9-116">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="1b6e9-116">
        - Custom Functions</span></span><br><span data-ttu-id="1b6e9-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="1b6e9-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="1b6e9-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1b6e9-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1b6e9-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1b6e9-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1b6e9-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1b6e9-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="1b6e9-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="1b6e9-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="1b6e9-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="1b6e9-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-128">
        - BindingEvents</span></span><br><span data-ttu-id="1b6e9-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-129">
        - CompressedFile</span></span><br><span data-ttu-id="1b6e9-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-130">
        - DocumentEvents</span></span><br><span data-ttu-id="1b6e9-131">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-131">
        - File</span></span><br><span data-ttu-id="1b6e9-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-132">
        - MatrixBindings</span></span><br><span data-ttu-id="1b6e9-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-133">
        - MatrixCoercion</span></span><br><span data-ttu-id="1b6e9-134">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="1b6e9-134">
        - Selection</span></span><br><span data-ttu-id="1b6e9-135">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="1b6e9-135">
        - Settings</span></span><br><span data-ttu-id="1b6e9-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-136">
        - TableBindings</span></span><br><span data-ttu-id="1b6e9-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-137">
        - TableCoercion</span></span><br><span data-ttu-id="1b6e9-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-138">
        - TextBindings</span></span><br><span data-ttu-id="1b6e9-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-139">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-140">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="1b6e9-140">Office on Windows</span></span><br><span data-ttu-id="1b6e9-141">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-141">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="1b6e9-142">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="1b6e9-142">- TaskPane</span></span><br><span data-ttu-id="1b6e9-143">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-143">
        - Content</span></span><br><span data-ttu-id="1b6e9-144">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="1b6e9-144">
        - Custom Functions</span></span><br><span data-ttu-id="1b6e9-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="1b6e9-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="1b6e9-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1b6e9-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1b6e9-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1b6e9-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1b6e9-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1b6e9-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="1b6e9-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="1b6e9-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="1b6e9-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="1b6e9-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="1b6e9-158">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-158">
        - BindingEvents</span></span><br><span data-ttu-id="1b6e9-159">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-159">
        - CompressedFile</span></span><br><span data-ttu-id="1b6e9-160">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-160">
        - DocumentEvents</span></span><br><span data-ttu-id="1b6e9-161">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-161">
        - File</span></span><br><span data-ttu-id="1b6e9-162">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-162">
        - MatrixBindings</span></span><br><span data-ttu-id="1b6e9-163">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-163">
        - MatrixCoercion</span></span><br><span data-ttu-id="1b6e9-164">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="1b6e9-164">
        - Selection</span></span><br><span data-ttu-id="1b6e9-165">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="1b6e9-165">
        - Settings</span></span><br><span data-ttu-id="1b6e9-166">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-166">
        - TableBindings</span></span><br><span data-ttu-id="1b6e9-167">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-167">
        - TableCoercion</span></span><br><span data-ttu-id="1b6e9-168">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-168">
        - TextBindings</span></span><br><span data-ttu-id="1b6e9-169">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-169">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-170">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="1b6e9-170">Office 2019 on Windows</span></span><br><span data-ttu-id="1b6e9-171">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-171">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="1b6e9-172">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="1b6e9-172">- TaskPane</span></span><br><span data-ttu-id="1b6e9-173">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-173">
        - Content</span></span><br><span data-ttu-id="1b6e9-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="1b6e9-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1b6e9-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1b6e9-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1b6e9-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1b6e9-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1b6e9-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="1b6e9-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="1b6e9-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="1b6e9-185">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-185">- BindingEvents</span></span><br><span data-ttu-id="1b6e9-186">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-186">
        - CompressedFile</span></span><br><span data-ttu-id="1b6e9-187">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-187">
        - DocumentEvents</span></span><br><span data-ttu-id="1b6e9-188">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-188">
        - File</span></span><br><span data-ttu-id="1b6e9-189">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-189">
        - MatrixBindings</span></span><br><span data-ttu-id="1b6e9-190">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-190">
        - MatrixCoercion</span></span><br><span data-ttu-id="1b6e9-191">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="1b6e9-191">
        - Selection</span></span><br><span data-ttu-id="1b6e9-192">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="1b6e9-192">
        - Settings</span></span><br><span data-ttu-id="1b6e9-193">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-193">
        - TableBindings</span></span><br><span data-ttu-id="1b6e9-194">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-194">
        - TableCoercion</span></span><br><span data-ttu-id="1b6e9-195">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-195">
        - TextBindings</span></span><br><span data-ttu-id="1b6e9-196">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-196">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-197">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="1b6e9-197">Office 2016 on Windows</span></span><br><span data-ttu-id="1b6e9-198">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-198">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="1b6e9-199">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="1b6e9-199">- TaskPane</span></span><br><span data-ttu-id="1b6e9-200">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-200">
        - Content</span></span></td>
    <td><span data-ttu-id="1b6e9-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="1b6e9-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="1b6e9-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="1b6e9-204">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-204">- BindingEvents</span></span><br><span data-ttu-id="1b6e9-205">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-205">
        - CompressedFile</span></span><br><span data-ttu-id="1b6e9-206">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-206">
        - DocumentEvents</span></span><br><span data-ttu-id="1b6e9-207">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-207">
        - File</span></span><br><span data-ttu-id="1b6e9-208">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-208">
        - MatrixBindings</span></span><br><span data-ttu-id="1b6e9-209">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-209">
        - MatrixCoercion</span></span><br><span data-ttu-id="1b6e9-210">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="1b6e9-210">
        - Selection</span></span><br><span data-ttu-id="1b6e9-211">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="1b6e9-211">
        - Settings</span></span><br><span data-ttu-id="1b6e9-212">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-212">
        - TableBindings</span></span><br><span data-ttu-id="1b6e9-213">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-213">
        - TableCoercion</span></span><br><span data-ttu-id="1b6e9-214">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-214">
        - TextBindings</span></span><br><span data-ttu-id="1b6e9-215">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-215">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-216">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="1b6e9-216">Office 2013 on Windows</span></span><br><span data-ttu-id="1b6e9-217">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-217">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="1b6e9-218">
        - TaskPane</span><span class="sxs-lookup"><span data-stu-id="1b6e9-218">
        - TaskPane</span></span><br><span data-ttu-id="1b6e9-219">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-219">
        - Content</span></span></td>
    <td>  <span data-ttu-id="1b6e9-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="1b6e9-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="1b6e9-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="1b6e9-222">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-222">
        - BindingEvents</span></span><br><span data-ttu-id="1b6e9-223">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-223">
        - CompressedFile</span></span><br><span data-ttu-id="1b6e9-224">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-224">
        - DocumentEvents</span></span><br><span data-ttu-id="1b6e9-225">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-225">
        - File</span></span><br><span data-ttu-id="1b6e9-226">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-226">
        - MatrixBindings</span></span><br><span data-ttu-id="1b6e9-227">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-227">
        - MatrixCoercion</span></span><br><span data-ttu-id="1b6e9-228">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="1b6e9-228">
        - Selection</span></span><br><span data-ttu-id="1b6e9-229">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="1b6e9-229">
        - Settings</span></span><br><span data-ttu-id="1b6e9-230">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-230">
        - TableBindings</span></span><br><span data-ttu-id="1b6e9-231">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-231">
        - TableCoercion</span></span><br><span data-ttu-id="1b6e9-232">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-232">
        - TextBindings</span></span><br><span data-ttu-id="1b6e9-233">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-233">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-234">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="1b6e9-234">Sideload Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="1b6e9-235">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-235">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="1b6e9-236">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="1b6e9-236">- TaskPane</span></span><br><span data-ttu-id="1b6e9-237">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-237">
        - Content</span></span></td>
    <td><span data-ttu-id="1b6e9-238">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-238">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1b6e9-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1b6e9-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1b6e9-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1b6e9-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1b6e9-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="1b6e9-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="1b6e9-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="1b6e9-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="1b6e9-249">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-249">- BindingEvents</span></span><br><span data-ttu-id="1b6e9-250">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-250">
        - DocumentEvents</span></span><br><span data-ttu-id="1b6e9-251">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-251">
        - File</span></span><br><span data-ttu-id="1b6e9-252">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-252">
        - MatrixBindings</span></span><br><span data-ttu-id="1b6e9-253">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-253">
        - MatrixCoercion</span></span><br><span data-ttu-id="1b6e9-254">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="1b6e9-254">
        - Selection</span></span><br><span data-ttu-id="1b6e9-255">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="1b6e9-255">
        - Settings</span></span><br><span data-ttu-id="1b6e9-256">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-256">
        - TableBindings</span></span><br><span data-ttu-id="1b6e9-257">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-257">
        - TableCoercion</span></span><br><span data-ttu-id="1b6e9-258">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-258">
        - TextBindings</span></span><br><span data-ttu-id="1b6e9-259">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-259">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-260">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="1b6e9-260">Office apps on Mac</span></span><br><span data-ttu-id="1b6e9-261">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-261">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="1b6e9-262">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="1b6e9-262">- TaskPane</span></span><br><span data-ttu-id="1b6e9-263">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-263">
        - Content</span></span><br><span data-ttu-id="1b6e9-264">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="1b6e9-264">
        - Custom Functions</span></span><br><span data-ttu-id="1b6e9-265">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-265">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="1b6e9-266">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-266">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1b6e9-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1b6e9-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1b6e9-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1b6e9-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1b6e9-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="1b6e9-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="1b6e9-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="1b6e9-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="1b6e9-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="1b6e9-278">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-278">- BindingEvents</span></span><br><span data-ttu-id="1b6e9-279">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-279">
        - CompressedFile</span></span><br><span data-ttu-id="1b6e9-280">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-280">
        - DocumentEvents</span></span><br><span data-ttu-id="1b6e9-281">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-281">
        - File</span></span><br><span data-ttu-id="1b6e9-282">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-282">
        - MatrixBindings</span></span><br><span data-ttu-id="1b6e9-283">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-283">
        - MatrixCoercion</span></span><br><span data-ttu-id="1b6e9-284">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-284">
        - PdfFile</span></span><br><span data-ttu-id="1b6e9-285">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="1b6e9-285">
        - Selection</span></span><br><span data-ttu-id="1b6e9-286">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="1b6e9-286">
        - Settings</span></span><br><span data-ttu-id="1b6e9-287">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-287">
        - TableBindings</span></span><br><span data-ttu-id="1b6e9-288">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-288">
        - TableCoercion</span></span><br><span data-ttu-id="1b6e9-289">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-289">
        - TextBindings</span></span><br><span data-ttu-id="1b6e9-290">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-290">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-291">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="1b6e9-291">Office 2019 for Mac</span></span><br><span data-ttu-id="1b6e9-292">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-292">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="1b6e9-293">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="1b6e9-293">- TaskPane</span></span><br><span data-ttu-id="1b6e9-294">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-294">
        - Content</span></span><br><span data-ttu-id="1b6e9-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="1b6e9-296">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-296">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1b6e9-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1b6e9-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1b6e9-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1b6e9-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1b6e9-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="1b6e9-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="1b6e9-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="1b6e9-306">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-306">- BindingEvents</span></span><br><span data-ttu-id="1b6e9-307">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-307">
        - CompressedFile</span></span><br><span data-ttu-id="1b6e9-308">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-308">
        - DocumentEvents</span></span><br><span data-ttu-id="1b6e9-309">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-309">
        - File</span></span><br><span data-ttu-id="1b6e9-310">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-310">
        - MatrixBindings</span></span><br><span data-ttu-id="1b6e9-311">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-311">
        - MatrixCoercion</span></span><br><span data-ttu-id="1b6e9-312">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-312">
        - PdfFile</span></span><br><span data-ttu-id="1b6e9-313">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="1b6e9-313">
        - Selection</span></span><br><span data-ttu-id="1b6e9-314">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="1b6e9-314">
        - Settings</span></span><br><span data-ttu-id="1b6e9-315">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-315">
        - TableBindings</span></span><br><span data-ttu-id="1b6e9-316">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-316">
        - TableCoercion</span></span><br><span data-ttu-id="1b6e9-317">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-317">
        - TextBindings</span></span><br><span data-ttu-id="1b6e9-318">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-318">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-319">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="1b6e9-319">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="1b6e9-320">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-320">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="1b6e9-321">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="1b6e9-321">- TaskPane</span></span><br><span data-ttu-id="1b6e9-322">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-322">
        - Content</span></span></td>
    <td><span data-ttu-id="1b6e9-323">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-323">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-324">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="1b6e9-324">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="1b6e9-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="1b6e9-326">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-326">- BindingEvents</span></span><br><span data-ttu-id="1b6e9-327">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-327">
        - CompressedFile</span></span><br><span data-ttu-id="1b6e9-328">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-328">
        - DocumentEvents</span></span><br><span data-ttu-id="1b6e9-329">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-329">
        - File</span></span><br><span data-ttu-id="1b6e9-330">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-330">
        - MatrixBindings</span></span><br><span data-ttu-id="1b6e9-331">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-331">
        - MatrixCoercion</span></span><br><span data-ttu-id="1b6e9-332">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-332">
        - PdfFile</span></span><br><span data-ttu-id="1b6e9-333">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="1b6e9-333">
        - Selection</span></span><br><span data-ttu-id="1b6e9-334">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="1b6e9-334">
        - Settings</span></span><br><span data-ttu-id="1b6e9-335">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-335">
        - TableBindings</span></span><br><span data-ttu-id="1b6e9-336">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-336">
        - TableCoercion</span></span><br><span data-ttu-id="1b6e9-337">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-337">
        - TextBindings</span></span><br><span data-ttu-id="1b6e9-338">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-338">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="1b6e9-339">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="1b6e9-339">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="1b6e9-340">Funções Personalizadas</span><span class="sxs-lookup"><span data-stu-id="1b6e9-340">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="1b6e9-341">Plataforma</span><span class="sxs-lookup"><span data-stu-id="1b6e9-341">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="1b6e9-342">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="1b6e9-342">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="1b6e9-343">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="1b6e9-343">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="1b6e9-344"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-344"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-345">Office na Web</span><span class="sxs-lookup"><span data-stu-id="1b6e9-345">Office on the web</span></span></td>
    <td><span data-ttu-id="1b6e9-346">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="1b6e9-346">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="1b6e9-347">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-347">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-348">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="1b6e9-348">Office on Windows</span></span><br><span data-ttu-id="1b6e9-349">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-349">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="1b6e9-350">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="1b6e9-350">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="1b6e9-351">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-351">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-352">Office para Mac</span><span class="sxs-lookup"><span data-stu-id="1b6e9-352">Office for Mac</span></span><br><span data-ttu-id="1b6e9-353">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-353">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="1b6e9-354">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="1b6e9-354">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="1b6e9-355">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-355">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="1b6e9-356">Outlook</span><span class="sxs-lookup"><span data-stu-id="1b6e9-356">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="1b6e9-357">Plataforma</span><span class="sxs-lookup"><span data-stu-id="1b6e9-357">Platform</span></span></th>
    <th><span data-ttu-id="1b6e9-358">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="1b6e9-358">Extension points</span></span></th>
    <th><span data-ttu-id="1b6e9-359">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="1b6e9-359">API requirement sets</span></span></th>
    <th><span data-ttu-id="1b6e9-360"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-360"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-361">Office na Web</span><span class="sxs-lookup"><span data-stu-id="1b6e9-361">Office on the web</span></span><br><span data-ttu-id="1b6e9-362">(moderno)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-362">Modern</span></span></td>
    <td> <span data-ttu-id="1b6e9-363">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="1b6e9-363">- Mail Read</span></span><br><span data-ttu-id="1b6e9-364">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="1b6e9-364">
      - Mail Compose</span></span><br><span data-ttu-id="1b6e9-365">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-365">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1b6e9-366">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-366">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1b6e9-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1b6e9-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1b6e9-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1b6e9-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1b6e9-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="1b6e9-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="1b6e9-373">Não disponível</span><span class="sxs-lookup"><span data-stu-id="1b6e9-373">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-374">Office na Web</span><span class="sxs-lookup"><span data-stu-id="1b6e9-374">Office on the web</span></span><br><span data-ttu-id="1b6e9-375">(clássico)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-375">Classic.</span></span></td>
    <td> <span data-ttu-id="1b6e9-376">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="1b6e9-376">- Mail Read</span></span><br><span data-ttu-id="1b6e9-377">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="1b6e9-377">
      - Mail Compose</span></span><br><span data-ttu-id="1b6e9-378">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-378">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1b6e9-379">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-379">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1b6e9-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1b6e9-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1b6e9-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1b6e9-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1b6e9-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="1b6e9-385">Não disponível</span><span class="sxs-lookup"><span data-stu-id="1b6e9-385">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-386">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="1b6e9-386">Office on Windows</span></span><br><span data-ttu-id="1b6e9-387">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-387">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="1b6e9-388">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="1b6e9-388">- Mail Read</span></span><br><span data-ttu-id="1b6e9-389">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="1b6e9-389">
      - Mail Compose</span></span><br><span data-ttu-id="1b6e9-390">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-390">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="1b6e9-391">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="1b6e9-391">
      - Modules</span></span></td>
    <td> <span data-ttu-id="1b6e9-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1b6e9-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1b6e9-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1b6e9-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1b6e9-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1b6e9-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="1b6e9-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="1b6e9-399">Não disponível</span><span class="sxs-lookup"><span data-stu-id="1b6e9-399">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-400">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="1b6e9-400">Office 2019 on Windows</span></span><br><span data-ttu-id="1b6e9-401">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-401">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1b6e9-402">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="1b6e9-402">- Mail Read</span></span><br><span data-ttu-id="1b6e9-403">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="1b6e9-403">
      - Mail Compose</span></span><br><span data-ttu-id="1b6e9-404">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-404">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="1b6e9-405">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="1b6e9-405">
      - Modules</span></span></td>
    <td> <span data-ttu-id="1b6e9-406">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-406">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1b6e9-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1b6e9-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1b6e9-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1b6e9-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1b6e9-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="1b6e9-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="1b6e9-413">Não disponível</span><span class="sxs-lookup"><span data-stu-id="1b6e9-413">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-414">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="1b6e9-414">Office 2016 on Windows</span></span><br><span data-ttu-id="1b6e9-415">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-415">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1b6e9-416">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="1b6e9-416">- Mail Read</span></span><br><span data-ttu-id="1b6e9-417">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="1b6e9-417">
      - Mail Compose</span></span><br><span data-ttu-id="1b6e9-418">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-418">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="1b6e9-419">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="1b6e9-419">
      - Modules</span></span></td>
    <td> <span data-ttu-id="1b6e9-420">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-420">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1b6e9-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1b6e9-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1b6e9-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="1b6e9-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="1b6e9-424">Não disponível</span><span class="sxs-lookup"><span data-stu-id="1b6e9-424">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-425">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="1b6e9-425">Office 2013 on Windows</span></span><br><span data-ttu-id="1b6e9-426">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-426">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1b6e9-427">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="1b6e9-427">- Mail Read</span></span><br><span data-ttu-id="1b6e9-428">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="1b6e9-428">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="1b6e9-429">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-429">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1b6e9-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1b6e9-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="1b6e9-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="1b6e9-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="1b6e9-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="1b6e9-433">Não disponível</span><span class="sxs-lookup"><span data-stu-id="1b6e9-433">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-434">Office no iOS</span><span class="sxs-lookup"><span data-stu-id="1b6e9-434">Office apps on iOS</span></span><br><span data-ttu-id="1b6e9-435">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-435">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="1b6e9-436">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="1b6e9-436">- Mail Read</span></span><br><span data-ttu-id="1b6e9-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1b6e9-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1b6e9-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1b6e9-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1b6e9-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1b6e9-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="1b6e9-443">Não disponível</span><span class="sxs-lookup"><span data-stu-id="1b6e9-443">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-444">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="1b6e9-444">Office apps on Mac</span></span><br><span data-ttu-id="1b6e9-445">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-445">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="1b6e9-446">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="1b6e9-446">- Mail Read</span></span><br><span data-ttu-id="1b6e9-447">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="1b6e9-447">
      - Mail Compose</span></span><br><span data-ttu-id="1b6e9-448">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-448">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1b6e9-449">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-449">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1b6e9-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1b6e9-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1b6e9-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1b6e9-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1b6e9-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="1b6e9-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="1b6e9-456">Não disponível</span><span class="sxs-lookup"><span data-stu-id="1b6e9-456">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-457">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="1b6e9-457">Office 2019 for Mac</span></span><br><span data-ttu-id="1b6e9-458">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-458">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1b6e9-459">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="1b6e9-459">- Mail Read</span></span><br><span data-ttu-id="1b6e9-460">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="1b6e9-460">
      - Mail Compose</span></span><br><span data-ttu-id="1b6e9-461">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-461">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1b6e9-462">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-462">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1b6e9-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1b6e9-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1b6e9-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1b6e9-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1b6e9-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="1b6e9-468">Não disponível</span><span class="sxs-lookup"><span data-stu-id="1b6e9-468">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-469">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="1b6e9-469">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="1b6e9-470">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-470">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1b6e9-471">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="1b6e9-471">- Mail Read</span></span><br><span data-ttu-id="1b6e9-472">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="1b6e9-472">
      - Mail Compose</span></span><br><span data-ttu-id="1b6e9-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1b6e9-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1b6e9-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1b6e9-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1b6e9-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1b6e9-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1b6e9-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="1b6e9-480">Não disponível</span><span class="sxs-lookup"><span data-stu-id="1b6e9-480">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-481">Outlook no Android</span><span class="sxs-lookup"><span data-stu-id="1b6e9-481">Office apps on Android</span></span><br><span data-ttu-id="1b6e9-482">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-482">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="1b6e9-483">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="1b6e9-483">- Mail Read</span></span><br><span data-ttu-id="1b6e9-484">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-484">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1b6e9-485">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-485">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1b6e9-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1b6e9-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1b6e9-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1b6e9-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="1b6e9-490">Não disponível</span><span class="sxs-lookup"><span data-stu-id="1b6e9-490">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="1b6e9-491">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="1b6e9-491">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="1b6e9-492">O suporte ao cliente para um conjunto de requisitos pode ser restringido pelo suporte do servidor Exchange.</span><span class="sxs-lookup"><span data-stu-id="1b6e9-492">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="1b6e9-493">Consulte [Conjuntos de requisitos da API JavaScript do Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) para obter detalhes sobre o intervalo de conjuntos de requisitos suportado pelo servidor Exchange e pelos clientes Outlook.</span><span class="sxs-lookup"><span data-stu-id="1b6e9-493">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="1b6e9-494">Word</span><span class="sxs-lookup"><span data-stu-id="1b6e9-494">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="1b6e9-495">Plataforma</span><span class="sxs-lookup"><span data-stu-id="1b6e9-495">Platform</span></span></th>
    <th><span data-ttu-id="1b6e9-496">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="1b6e9-496">Extension points</span></span></th>
    <th><span data-ttu-id="1b6e9-497">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="1b6e9-497">API requirement sets</span></span></th>
    <th><span data-ttu-id="1b6e9-498"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-498"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-499">Office na Web</span><span class="sxs-lookup"><span data-stu-id="1b6e9-499">Office on the web</span></span></td>
    <td> <span data-ttu-id="1b6e9-500">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="1b6e9-500">- TaskPane</span></span><br><span data-ttu-id="1b6e9-501">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-501">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1b6e9-502">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-502">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-503">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-503">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="1b6e9-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="1b6e9-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="1b6e9-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="1b6e9-508">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-508">- BindingEvents</span></span><br><span data-ttu-id="1b6e9-509">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1b6e9-509">
         - CustomXmlParts</span></span><br><span data-ttu-id="1b6e9-510">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-510">
         - DocumentEvents</span></span><br><span data-ttu-id="1b6e9-511">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-511">
         - File</span></span><br><span data-ttu-id="1b6e9-512">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-512">
         - HtmlCoercion</span></span><br><span data-ttu-id="1b6e9-513">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-513">
         - MatrixBindings</span></span><br><span data-ttu-id="1b6e9-514">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-514">
         - MatrixCoercion</span></span><br><span data-ttu-id="1b6e9-515">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-515">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1b6e9-516">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-516">
         - PdfFile</span></span><br><span data-ttu-id="1b6e9-517">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="1b6e9-517">
         - Selection</span></span><br><span data-ttu-id="1b6e9-518">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="1b6e9-518">
         - Settings</span></span><br><span data-ttu-id="1b6e9-519">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-519">
         - TableBindings</span></span><br><span data-ttu-id="1b6e9-520">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-520">
         - TableCoercion</span></span><br><span data-ttu-id="1b6e9-521">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-521">
         - TextBindings</span></span><br><span data-ttu-id="1b6e9-522">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-522">
         - TextCoercion</span></span><br><span data-ttu-id="1b6e9-523">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-523">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-524">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="1b6e9-524">Office on Windows</span></span><br><span data-ttu-id="1b6e9-525">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-525">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="1b6e9-526">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="1b6e9-526">- TaskPane</span></span><br><span data-ttu-id="1b6e9-527">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-527">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1b6e9-528">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-528">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-529">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-529">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="1b6e9-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="1b6e9-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-532">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-532">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="1b6e9-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="1b6e9-534">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-534">- BindingEvents</span></span><br><span data-ttu-id="1b6e9-535">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-535">
         - CompressedFile</span></span><br><span data-ttu-id="1b6e9-536">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1b6e9-536">
         - CustomXmlParts</span></span><br><span data-ttu-id="1b6e9-537">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-537">
         - DocumentEvents</span></span><br><span data-ttu-id="1b6e9-538">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-538">
         - File</span></span><br><span data-ttu-id="1b6e9-539">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-539">
         - HtmlCoercion</span></span><br><span data-ttu-id="1b6e9-540">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-540">
         - MatrixBindings</span></span><br><span data-ttu-id="1b6e9-541">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-541">
         - MatrixCoercion</span></span><br><span data-ttu-id="1b6e9-542">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-542">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1b6e9-543">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-543">
         - PdfFile</span></span><br><span data-ttu-id="1b6e9-544">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="1b6e9-544">
         - Selection</span></span><br><span data-ttu-id="1b6e9-545">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="1b6e9-545">
         - Settings</span></span><br><span data-ttu-id="1b6e9-546">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-546">
         - TableBindings</span></span><br><span data-ttu-id="1b6e9-547">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-547">
         - TableCoercion</span></span><br><span data-ttu-id="1b6e9-548">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-548">
         - TextBindings</span></span><br><span data-ttu-id="1b6e9-549">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-549">
         - TextCoercion</span></span><br><span data-ttu-id="1b6e9-550">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-550">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-551">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="1b6e9-551">Office 2019 on Windows</span></span><br><span data-ttu-id="1b6e9-552">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-552">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1b6e9-553">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="1b6e9-553">- TaskPane</span></span><br><span data-ttu-id="1b6e9-554">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-554">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1b6e9-555">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-555">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="1b6e9-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="1b6e9-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-559">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-559">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="1b6e9-560">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-560">- BindingEvents</span></span><br><span data-ttu-id="1b6e9-561">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-561">
         - CompressedFile</span></span><br><span data-ttu-id="1b6e9-562">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1b6e9-562">
         - CustomXmlParts</span></span><br><span data-ttu-id="1b6e9-563">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-563">
         - DocumentEvents</span></span><br><span data-ttu-id="1b6e9-564">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-564">
         - File</span></span><br><span data-ttu-id="1b6e9-565">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-565">
         - HtmlCoercion</span></span><br><span data-ttu-id="1b6e9-566">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-566">
         - MatrixBindings</span></span><br><span data-ttu-id="1b6e9-567">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-567">
         - MatrixCoercion</span></span><br><span data-ttu-id="1b6e9-568">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-568">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1b6e9-569">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-569">
         - PdfFile</span></span><br><span data-ttu-id="1b6e9-570">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="1b6e9-570">
         - Selection</span></span><br><span data-ttu-id="1b6e9-571">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="1b6e9-571">
         - Settings</span></span><br><span data-ttu-id="1b6e9-572">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-572">
         - TableBindings</span></span><br><span data-ttu-id="1b6e9-573">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-573">
         - TableCoercion</span></span><br><span data-ttu-id="1b6e9-574">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-574">
         - TextBindings</span></span><br><span data-ttu-id="1b6e9-575">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-575">
         - TextCoercion</span></span><br><span data-ttu-id="1b6e9-576">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-576">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-577">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="1b6e9-577">Office 2016 on Windows</span></span><br><span data-ttu-id="1b6e9-578">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-578">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1b6e9-579">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="1b6e9-579">- TaskPane</span></span></td>
    <td> <span data-ttu-id="1b6e9-580">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-580">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-581">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="1b6e9-581">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="1b6e9-582">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-582">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="1b6e9-583">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-583">- BindingEvents</span></span><br><span data-ttu-id="1b6e9-584">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-584">
         - CompressedFile</span></span><br><span data-ttu-id="1b6e9-585">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1b6e9-585">
         - CustomXmlParts</span></span><br><span data-ttu-id="1b6e9-586">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-586">
         - DocumentEvents</span></span><br><span data-ttu-id="1b6e9-587">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-587">
         - File</span></span><br><span data-ttu-id="1b6e9-588">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-588">
         - HtmlCoercion</span></span><br><span data-ttu-id="1b6e9-589">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-589">
         - MatrixBindings</span></span><br><span data-ttu-id="1b6e9-590">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-590">
         - MatrixCoercion</span></span><br><span data-ttu-id="1b6e9-591">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-591">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1b6e9-592">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-592">
         - PdfFile</span></span><br><span data-ttu-id="1b6e9-593">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="1b6e9-593">
         - Selection</span></span><br><span data-ttu-id="1b6e9-594">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="1b6e9-594">
         - Settings</span></span><br><span data-ttu-id="1b6e9-595">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-595">
         - TableBindings</span></span><br><span data-ttu-id="1b6e9-596">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-596">
         - TableCoercion</span></span><br><span data-ttu-id="1b6e9-597">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-597">
         - TextBindings</span></span><br><span data-ttu-id="1b6e9-598">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-598">
         - TextCoercion</span></span><br><span data-ttu-id="1b6e9-599">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-599">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-600">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="1b6e9-600">Office 2013 on Windows</span></span><br><span data-ttu-id="1b6e9-601">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-601">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1b6e9-602">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="1b6e9-602">- TaskPane</span></span></td>
    <td> <span data-ttu-id="1b6e9-603">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="1b6e9-603">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="1b6e9-604">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-604">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="1b6e9-605">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-605">- BindingEvents</span></span><br><span data-ttu-id="1b6e9-606">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-606">
         - CompressedFile</span></span><br><span data-ttu-id="1b6e9-607">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1b6e9-607">
         - CustomXmlParts</span></span><br><span data-ttu-id="1b6e9-608">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-608">
         - DocumentEvents</span></span><br><span data-ttu-id="1b6e9-609">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-609">
         - File</span></span><br><span data-ttu-id="1b6e9-610">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-610">
         - HtmlCoercion</span></span><br><span data-ttu-id="1b6e9-611">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-611">
         - MatrixBindings</span></span><br><span data-ttu-id="1b6e9-612">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-612">
         - MatrixCoercion</span></span><br><span data-ttu-id="1b6e9-613">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-613">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1b6e9-614">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-614">
         - PdfFile</span></span><br><span data-ttu-id="1b6e9-615">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="1b6e9-615">
         - Selection</span></span><br><span data-ttu-id="1b6e9-616">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="1b6e9-616">
         - Settings</span></span><br><span data-ttu-id="1b6e9-617">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-617">
         - TableBindings</span></span><br><span data-ttu-id="1b6e9-618">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-618">
         - TableCoercion</span></span><br><span data-ttu-id="1b6e9-619">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-619">
         - TextBindings</span></span><br><span data-ttu-id="1b6e9-620">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-620">
         - TextCoercion</span></span><br><span data-ttu-id="1b6e9-621">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-621">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-622">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="1b6e9-622">Sideload Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="1b6e9-623">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-623">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="1b6e9-624">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="1b6e9-624">- TaskPane</span></span></td>
    <td> <span data-ttu-id="1b6e9-625">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-625">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-626">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-626">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="1b6e9-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="1b6e9-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="1b6e9-630">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-630">- BindingEvents</span></span><br><span data-ttu-id="1b6e9-631">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-631">
         - CompressedFile</span></span><br><span data-ttu-id="1b6e9-632">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1b6e9-632">
         - CustomXmlParts</span></span><br><span data-ttu-id="1b6e9-633">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-633">
         - DocumentEvents</span></span><br><span data-ttu-id="1b6e9-634">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-634">
         - File</span></span><br><span data-ttu-id="1b6e9-635">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-635">
         - HtmlCoercion</span></span><br><span data-ttu-id="1b6e9-636">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-636">
         - MatrixBindings</span></span><br><span data-ttu-id="1b6e9-637">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-637">
         - MatrixCoercion</span></span><br><span data-ttu-id="1b6e9-638">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-638">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1b6e9-639">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-639">
         - PdfFile</span></span><br><span data-ttu-id="1b6e9-640">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="1b6e9-640">
         - Selection</span></span><br><span data-ttu-id="1b6e9-641">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="1b6e9-641">
         - Settings</span></span><br><span data-ttu-id="1b6e9-642">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-642">
         - TableBindings</span></span><br><span data-ttu-id="1b6e9-643">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-643">
         - TableCoercion</span></span><br><span data-ttu-id="1b6e9-644">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-644">
         - TextBindings</span></span><br><span data-ttu-id="1b6e9-645">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-645">
         - TextCoercion</span></span><br><span data-ttu-id="1b6e9-646">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-646">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-647">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="1b6e9-647">Office apps on Mac</span></span><br><span data-ttu-id="1b6e9-648">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-648">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="1b6e9-649">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="1b6e9-649">- TaskPane</span></span><br><span data-ttu-id="1b6e9-650">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-650">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1b6e9-651">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-651">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-652">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-652">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="1b6e9-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="1b6e9-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="1b6e9-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="1b6e9-657">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-657">- BindingEvents</span></span><br><span data-ttu-id="1b6e9-658">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-658">
         - CompressedFile</span></span><br><span data-ttu-id="1b6e9-659">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1b6e9-659">
         - CustomXmlParts</span></span><br><span data-ttu-id="1b6e9-660">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-660">
         - DocumentEvents</span></span><br><span data-ttu-id="1b6e9-661">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-661">
         - File</span></span><br><span data-ttu-id="1b6e9-662">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-662">
         - HtmlCoercion</span></span><br><span data-ttu-id="1b6e9-663">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-663">
         - MatrixBindings</span></span><br><span data-ttu-id="1b6e9-664">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-664">
         - MatrixCoercion</span></span><br><span data-ttu-id="1b6e9-665">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-665">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1b6e9-666">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-666">
         - PdfFile</span></span><br><span data-ttu-id="1b6e9-667">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="1b6e9-667">
         - Selection</span></span><br><span data-ttu-id="1b6e9-668">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="1b6e9-668">
         - Settings</span></span><br><span data-ttu-id="1b6e9-669">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-669">
         - TableBindings</span></span><br><span data-ttu-id="1b6e9-670">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-670">
         - TableCoercion</span></span><br><span data-ttu-id="1b6e9-671">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-671">
         - TextBindings</span></span><br><span data-ttu-id="1b6e9-672">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-672">
         - TextCoercion</span></span><br><span data-ttu-id="1b6e9-673">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-673">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-674">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="1b6e9-674">Office 2019 for Mac</span></span><br><span data-ttu-id="1b6e9-675">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-675">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1b6e9-676">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="1b6e9-676">- TaskPane</span></span><br><span data-ttu-id="1b6e9-677">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-677">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1b6e9-678">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-678">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="1b6e9-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="1b6e9-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="1b6e9-683">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-683">- BindingEvents</span></span><br><span data-ttu-id="1b6e9-684">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-684">
         - CompressedFile</span></span><br><span data-ttu-id="1b6e9-685">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1b6e9-685">
         - CustomXmlParts</span></span><br><span data-ttu-id="1b6e9-686">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-686">
         - DocumentEvents</span></span><br><span data-ttu-id="1b6e9-687">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-687">
         - File</span></span><br><span data-ttu-id="1b6e9-688">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-688">
         - HtmlCoercion</span></span><br><span data-ttu-id="1b6e9-689">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-689">
         - MatrixBindings</span></span><br><span data-ttu-id="1b6e9-690">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-690">
         - MatrixCoercion</span></span><br><span data-ttu-id="1b6e9-691">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-691">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1b6e9-692">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-692">
         - PdfFile</span></span><br><span data-ttu-id="1b6e9-693">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="1b6e9-693">
         - Selection</span></span><br><span data-ttu-id="1b6e9-694">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="1b6e9-694">
         - Settings</span></span><br><span data-ttu-id="1b6e9-695">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-695">
         - TableBindings</span></span><br><span data-ttu-id="1b6e9-696">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-696">
         - TableCoercion</span></span><br><span data-ttu-id="1b6e9-697">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-697">
         - TextBindings</span></span><br><span data-ttu-id="1b6e9-698">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-698">
         - TextCoercion</span></span><br><span data-ttu-id="1b6e9-699">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-699">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-700">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="1b6e9-700">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="1b6e9-701">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-701">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1b6e9-702">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="1b6e9-702">- TaskPane</span></span></td>
    <td> <span data-ttu-id="1b6e9-703">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-703">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-704">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="1b6e9-704">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="1b6e9-705">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-705">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="1b6e9-706">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-706">- BindingEvents</span></span><br><span data-ttu-id="1b6e9-707">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-707">
         - CompressedFile</span></span><br><span data-ttu-id="1b6e9-708">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1b6e9-708">
         - CustomXmlParts</span></span><br><span data-ttu-id="1b6e9-709">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-709">
         - DocumentEvents</span></span><br><span data-ttu-id="1b6e9-710">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-710">
         - File</span></span><br><span data-ttu-id="1b6e9-711">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-711">
         - HtmlCoercion</span></span><br><span data-ttu-id="1b6e9-712">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-712">
         - MatrixBindings</span></span><br><span data-ttu-id="1b6e9-713">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-713">
         - MatrixCoercion</span></span><br><span data-ttu-id="1b6e9-714">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-714">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1b6e9-715">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-715">
         - PdfFile</span></span><br><span data-ttu-id="1b6e9-716">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="1b6e9-716">
         - Selection</span></span><br><span data-ttu-id="1b6e9-717">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="1b6e9-717">
         - Settings</span></span><br><span data-ttu-id="1b6e9-718">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-718">
         - TableBindings</span></span><br><span data-ttu-id="1b6e9-719">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-719">
         - TableCoercion</span></span><br><span data-ttu-id="1b6e9-720">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1b6e9-720">
         - TextBindings</span></span><br><span data-ttu-id="1b6e9-721">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-721">
         - TextCoercion</span></span><br><span data-ttu-id="1b6e9-722">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-722">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="1b6e9-723">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="1b6e9-723">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="1b6e9-724">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="1b6e9-724">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="1b6e9-725">Plataforma</span><span class="sxs-lookup"><span data-stu-id="1b6e9-725">Platform</span></span></th>
    <th><span data-ttu-id="1b6e9-726">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="1b6e9-726">Extension points</span></span></th>
    <th><span data-ttu-id="1b6e9-727">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="1b6e9-727">API requirement sets</span></span></th>
    <th><span data-ttu-id="1b6e9-728"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-728"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-729">Office na Web</span><span class="sxs-lookup"><span data-stu-id="1b6e9-729">Office on the web</span></span></td>
    <td> <span data-ttu-id="1b6e9-730">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-730">- Content</span></span><br><span data-ttu-id="1b6e9-731">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="1b6e9-731">
         - TaskPane</span></span><br><span data-ttu-id="1b6e9-732">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-732">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1b6e9-733">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-733">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-734">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-734">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="1b6e9-736">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-736">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="1b6e9-737">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1b6e9-737">- ActiveView</span></span><br><span data-ttu-id="1b6e9-738">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-738">
         - CompressedFile</span></span><br><span data-ttu-id="1b6e9-739">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-739">
         - DocumentEvents</span></span><br><span data-ttu-id="1b6e9-740">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-740">
         - File</span></span><br><span data-ttu-id="1b6e9-741">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-741">
         - PdfFile</span></span><br><span data-ttu-id="1b6e9-742">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="1b6e9-742">
         - Selection</span></span><br><span data-ttu-id="1b6e9-743">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="1b6e9-743">
         - Settings</span></span><br><span data-ttu-id="1b6e9-744">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-744">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-745">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="1b6e9-745">Office on Windows</span></span><br><span data-ttu-id="1b6e9-746">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-746">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="1b6e9-747">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-747">- Content</span></span><br><span data-ttu-id="1b6e9-748">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="1b6e9-748">
         - TaskPane</span></span><br><span data-ttu-id="1b6e9-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1b6e9-750">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-750">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="1b6e9-753">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-753">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="1b6e9-754">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1b6e9-754">- ActiveView</span></span><br><span data-ttu-id="1b6e9-755">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-755">
         - CompressedFile</span></span><br><span data-ttu-id="1b6e9-756">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-756">
         - DocumentEvents</span></span><br><span data-ttu-id="1b6e9-757">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-757">
         - File</span></span><br><span data-ttu-id="1b6e9-758">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-758">
         - PdfFile</span></span><br><span data-ttu-id="1b6e9-759">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="1b6e9-759">
         - Selection</span></span><br><span data-ttu-id="1b6e9-760">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="1b6e9-760">
         - Settings</span></span><br><span data-ttu-id="1b6e9-761">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-761">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-762">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="1b6e9-762">Office 2019 on Windows</span></span><br><span data-ttu-id="1b6e9-763">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-763">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1b6e9-764">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-764">- Content</span></span><br><span data-ttu-id="1b6e9-765">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="1b6e9-765">
         - TaskPane</span></span><br><span data-ttu-id="1b6e9-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1b6e9-767">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-767">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-768">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-768">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="1b6e9-769">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1b6e9-769">- ActiveView</span></span><br><span data-ttu-id="1b6e9-770">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-770">
         - CompressedFile</span></span><br><span data-ttu-id="1b6e9-771">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-771">
         - DocumentEvents</span></span><br><span data-ttu-id="1b6e9-772">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-772">
         - File</span></span><br><span data-ttu-id="1b6e9-773">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-773">
         - PdfFile</span></span><br><span data-ttu-id="1b6e9-774">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="1b6e9-774">
         - Selection</span></span><br><span data-ttu-id="1b6e9-775">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="1b6e9-775">
         - Settings</span></span><br><span data-ttu-id="1b6e9-776">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-776">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-777">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="1b6e9-777">Office 2016 on Windows</span></span><br><span data-ttu-id="1b6e9-778">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-778">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1b6e9-779">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-779">- Content</span></span><br><span data-ttu-id="1b6e9-780">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="1b6e9-780">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="1b6e9-781">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="1b6e9-781">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="1b6e9-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="1b6e9-783">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1b6e9-783">- ActiveView</span></span><br><span data-ttu-id="1b6e9-784">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-784">
         - CompressedFile</span></span><br><span data-ttu-id="1b6e9-785">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-785">
         - DocumentEvents</span></span><br><span data-ttu-id="1b6e9-786">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-786">
         - File</span></span><br><span data-ttu-id="1b6e9-787">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-787">
         - PdfFile</span></span><br><span data-ttu-id="1b6e9-788">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="1b6e9-788">
         - Selection</span></span><br><span data-ttu-id="1b6e9-789">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="1b6e9-789">
         - Settings</span></span><br><span data-ttu-id="1b6e9-790">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-790">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-791">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="1b6e9-791">Office 2013 on Windows</span></span><br><span data-ttu-id="1b6e9-792">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-792">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1b6e9-793">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-793">- Content</span></span><br><span data-ttu-id="1b6e9-794">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="1b6e9-794">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="1b6e9-795">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="1b6e9-795">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="1b6e9-796">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-796">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="1b6e9-797">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1b6e9-797">- ActiveView</span></span><br><span data-ttu-id="1b6e9-798">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-798">
         - CompressedFile</span></span><br><span data-ttu-id="1b6e9-799">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-799">
         - DocumentEvents</span></span><br><span data-ttu-id="1b6e9-800">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-800">
         - File</span></span><br><span data-ttu-id="1b6e9-801">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-801">
         - PdfFile</span></span><br><span data-ttu-id="1b6e9-802">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="1b6e9-802">
         - Selection</span></span><br><span data-ttu-id="1b6e9-803">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="1b6e9-803">
         - Settings</span></span><br><span data-ttu-id="1b6e9-804">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-804">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-805">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="1b6e9-805">Sideload Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="1b6e9-806">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-806">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="1b6e9-807">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-807">- Content</span></span><br><span data-ttu-id="1b6e9-808">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="1b6e9-808">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="1b6e9-809">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-809">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-810">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-810">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-811">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-811">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="1b6e9-812">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1b6e9-812">- ActiveView</span></span><br><span data-ttu-id="1b6e9-813">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-813">
         - CompressedFile</span></span><br><span data-ttu-id="1b6e9-814">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-814">
         - DocumentEvents</span></span><br><span data-ttu-id="1b6e9-815">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-815">
         - File</span></span><br><span data-ttu-id="1b6e9-816">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-816">
         - PdfFile</span></span><br><span data-ttu-id="1b6e9-817">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="1b6e9-817">
         - Selection</span></span><br><span data-ttu-id="1b6e9-818">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="1b6e9-818">
         - Settings</span></span><br><span data-ttu-id="1b6e9-819">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-819">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-820">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="1b6e9-820">Office apps on Mac</span></span><br><span data-ttu-id="1b6e9-821">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-821">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="1b6e9-822">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-822">- Content</span></span><br><span data-ttu-id="1b6e9-823">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="1b6e9-823">
         - TaskPane</span></span><br><span data-ttu-id="1b6e9-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1b6e9-825">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-825">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="1b6e9-828">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-828">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="1b6e9-829">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1b6e9-829">- ActiveView</span></span><br><span data-ttu-id="1b6e9-830">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-830">
         - CompressedFile</span></span><br><span data-ttu-id="1b6e9-831">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-831">
         - DocumentEvents</span></span><br><span data-ttu-id="1b6e9-832">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-832">
         - File</span></span><br><span data-ttu-id="1b6e9-833">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-833">
         - PdfFile</span></span><br><span data-ttu-id="1b6e9-834">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="1b6e9-834">
         - Selection</span></span><br><span data-ttu-id="1b6e9-835">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="1b6e9-835">
         - Settings</span></span><br><span data-ttu-id="1b6e9-836">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-836">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-837">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="1b6e9-837">Office 2019 for Mac</span></span><br><span data-ttu-id="1b6e9-838">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-838">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1b6e9-839">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-839">- Content</span></span><br><span data-ttu-id="1b6e9-840">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="1b6e9-840">
         - TaskPane</span></span><br><span data-ttu-id="1b6e9-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1b6e9-842">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-842">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-843">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-843">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="1b6e9-844">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1b6e9-844">- ActiveView</span></span><br><span data-ttu-id="1b6e9-845">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-845">
         - CompressedFile</span></span><br><span data-ttu-id="1b6e9-846">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-846">
         - DocumentEvents</span></span><br><span data-ttu-id="1b6e9-847">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-847">
         - File</span></span><br><span data-ttu-id="1b6e9-848">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-848">
         - PdfFile</span></span><br><span data-ttu-id="1b6e9-849">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="1b6e9-849">
         - Selection</span></span><br><span data-ttu-id="1b6e9-850">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="1b6e9-850">
         - Settings</span></span><br><span data-ttu-id="1b6e9-851">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-851">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-852">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="1b6e9-852">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="1b6e9-853">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-853">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1b6e9-854">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-854">- Content</span></span><br><span data-ttu-id="1b6e9-855">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="1b6e9-855">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="1b6e9-856">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="1b6e9-856">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="1b6e9-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="1b6e9-858">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1b6e9-858">- ActiveView</span></span><br><span data-ttu-id="1b6e9-859">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-859">
         - CompressedFile</span></span><br><span data-ttu-id="1b6e9-860">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-860">
         - DocumentEvents</span></span><br><span data-ttu-id="1b6e9-861">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-861">
         - File</span></span><br><span data-ttu-id="1b6e9-862">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1b6e9-862">
         - PdfFile</span></span><br><span data-ttu-id="1b6e9-863">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="1b6e9-863">
         - Selection</span></span><br><span data-ttu-id="1b6e9-864">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="1b6e9-864">
         - Settings</span></span><br><span data-ttu-id="1b6e9-865">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-865">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="1b6e9-866">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="1b6e9-866">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="1b6e9-867">OneNote</span><span class="sxs-lookup"><span data-stu-id="1b6e9-867">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="1b6e9-868">Plataforma</span><span class="sxs-lookup"><span data-stu-id="1b6e9-868">Platform</span></span></th>
    <th><span data-ttu-id="1b6e9-869">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="1b6e9-869">Extension points</span></span></th>
    <th><span data-ttu-id="1b6e9-870">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="1b6e9-870">API requirement sets</span></span></th>
    <th><span data-ttu-id="1b6e9-871"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-871"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-872">Office na Web</span><span class="sxs-lookup"><span data-stu-id="1b6e9-872">Office on the web</span></span></td>
    <td> <span data-ttu-id="1b6e9-873">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="1b6e9-873">- Content</span></span><br><span data-ttu-id="1b6e9-874">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="1b6e9-874">
         - TaskPane</span></span><br><span data-ttu-id="1b6e9-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1b6e9-876">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-876">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-877">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-877">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="1b6e9-878">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-878">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="1b6e9-879">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1b6e9-879">- DocumentEvents</span></span><br><span data-ttu-id="1b6e9-880">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-880">
         - HtmlCoercion</span></span><br><span data-ttu-id="1b6e9-881">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="1b6e9-881">
         - Settings</span></span><br><span data-ttu-id="1b6e9-882">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-882">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="1b6e9-883">Project</span><span class="sxs-lookup"><span data-stu-id="1b6e9-883">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="1b6e9-884">Plataforma</span><span class="sxs-lookup"><span data-stu-id="1b6e9-884">Platform</span></span></th>
    <th><span data-ttu-id="1b6e9-885">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="1b6e9-885">Extension points</span></span></th>
    <th><span data-ttu-id="1b6e9-886">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="1b6e9-886">API requirement sets</span></span></th>
    <th><span data-ttu-id="1b6e9-887"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-887"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-888">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="1b6e9-888">Office 2019 on Windows</span></span><br><span data-ttu-id="1b6e9-889">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-889">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1b6e9-890">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="1b6e9-890">- TaskPane</span></span></td>
    <td> <span data-ttu-id="1b6e9-891">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-891">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1b6e9-892">- Seleção</span><span class="sxs-lookup"><span data-stu-id="1b6e9-892">- Selection</span></span><br><span data-ttu-id="1b6e9-893">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-893">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-894">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="1b6e9-894">Office 2016 on Windows</span></span><br><span data-ttu-id="1b6e9-895">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-895">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1b6e9-896">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="1b6e9-896">- TaskPane</span></span></td>
    <td> <span data-ttu-id="1b6e9-897">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-897">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1b6e9-898">- Seleção</span><span class="sxs-lookup"><span data-stu-id="1b6e9-898">- Selection</span></span><br><span data-ttu-id="1b6e9-899">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-899">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1b6e9-900">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="1b6e9-900">Office 2013 on Windows</span></span><br><span data-ttu-id="1b6e9-901">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-901">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1b6e9-902">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="1b6e9-902">- TaskPane</span></span></td>
    <td> <span data-ttu-id="1b6e9-903">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1b6e9-903">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1b6e9-904">- Seleção</span><span class="sxs-lookup"><span data-stu-id="1b6e9-904">- Selection</span></span><br><span data-ttu-id="1b6e9-905">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1b6e9-905">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="1b6e9-906">Confira também</span><span class="sxs-lookup"><span data-stu-id="1b6e9-906">See also</span></span>

- [<span data-ttu-id="1b6e9-907">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="1b6e9-907">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="1b6e9-908">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="1b6e9-908">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="1b6e9-909">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="1b6e9-909">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="1b6e9-910">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="1b6e9-910">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="1b6e9-911">Referência da API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="1b6e9-911">JavaScript API for Office reference</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="1b6e9-912">Histórico de atualizações do Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="1b6e9-912">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="1b6e9-913">Histórico de atualizações do Office 2016 e 2019 (Clique para Executar)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-913">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="1b6e9-914">Histórico de atualizações do Office 2013 (clique para executar)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-914">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="1b6e9-915">Histórico de atualizações do Office 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-915">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="1b6e9-916">Histórico de atualizações do Outlook 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="1b6e9-916">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="1b6e9-917">Histórico de atualizações do Office para Mac</span><span class="sxs-lookup"><span data-stu-id="1b6e9-917">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
