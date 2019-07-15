---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos com suporte para o Excel, OneNote, Outlook, PowerPoint, Project e Word.
ms.date: 07/11/2019
localization_priority: Priority
ms.openlocfilehash: d88f7c1b9daa201d9b6bc5cfa69ac3125bf127b1
ms.sourcegitcommit: 61f8f02193ce05da957418d938f0d94cb12c468d
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/11/2019
ms.locfileid: "35630533"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="38f8d-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="38f8d-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="38f8d-104">Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="38f8d-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="38f8d-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que são compatíveis atualmente com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="38f8d-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="38f8d-106">A versão inicial do Office 2016 instalada por meio do MSI apenas contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="38f8d-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="38f8d-107">Para saber mais sobre o histórico de atualizações de várias versões do Office, confira a seção[Confira também](#see-also).</span><span class="sxs-lookup"><span data-stu-id="38f8d-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="38f8d-108">Excel</span><span class="sxs-lookup"><span data-stu-id="38f8d-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="38f8d-109">Plataforma</span><span class="sxs-lookup"><span data-stu-id="38f8d-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="38f8d-110">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="38f8d-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="38f8d-111">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="38f8d-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="38f8d-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="38f8d-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-113">Office na Web</span><span class="sxs-lookup"><span data-stu-id="38f8d-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="38f8d-114">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="38f8d-114">- TaskPane</span></span><br><span data-ttu-id="38f8d-115">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="38f8d-115">
        - Content</span></span><br><span data-ttu-id="38f8d-116">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="38f8d-116">
        - Custom Functions</span></span><br><span data-ttu-id="38f8d-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="38f8d-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="38f8d-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="38f8d-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="38f8d-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="38f8d-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="38f8d-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="38f8d-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="38f8d-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="38f8d-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="38f8d-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="38f8d-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="38f8d-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="38f8d-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="38f8d-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-130">
        - BindingEvents</span></span><br><span data-ttu-id="38f8d-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-131">
        - CompressedFile</span></span><br><span data-ttu-id="38f8d-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-132">
        - DocumentEvents</span></span><br><span data-ttu-id="38f8d-133">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="38f8d-133">
        - File</span></span><br><span data-ttu-id="38f8d-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-134">
        - MatrixBindings</span></span><br><span data-ttu-id="38f8d-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="38f8d-136">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="38f8d-136">
        - Selection</span></span><br><span data-ttu-id="38f8d-137">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="38f8d-137">
        - Settings</span></span><br><span data-ttu-id="38f8d-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-138">
        - TableBindings</span></span><br><span data-ttu-id="38f8d-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-139">
        - TableCoercion</span></span><br><span data-ttu-id="38f8d-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-140">
        - TextBindings</span></span><br><span data-ttu-id="38f8d-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-142">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="38f8d-142">Office on Windows</span></span><br><span data-ttu-id="38f8d-143">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="38f8d-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="38f8d-144">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="38f8d-144">- TaskPane</span></span><br><span data-ttu-id="38f8d-145">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="38f8d-145">
        - Content</span></span><br><span data-ttu-id="38f8d-146">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="38f8d-146">
        - Custom Functions</span></span><br><span data-ttu-id="38f8d-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="38f8d-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="38f8d-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="38f8d-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="38f8d-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="38f8d-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="38f8d-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="38f8d-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="38f8d-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="38f8d-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="38f8d-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="38f8d-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="38f8d-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="38f8d-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="38f8d-160">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-160">
        - BindingEvents</span></span><br><span data-ttu-id="38f8d-161">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-161">
        - CompressedFile</span></span><br><span data-ttu-id="38f8d-162">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-162">
        - DocumentEvents</span></span><br><span data-ttu-id="38f8d-163">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="38f8d-163">
        - File</span></span><br><span data-ttu-id="38f8d-164">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-164">
        - MatrixBindings</span></span><br><span data-ttu-id="38f8d-165">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-165">
        - MatrixCoercion</span></span><br><span data-ttu-id="38f8d-166">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="38f8d-166">
        - Selection</span></span><br><span data-ttu-id="38f8d-167">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="38f8d-167">
        - Settings</span></span><br><span data-ttu-id="38f8d-168">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-168">
        - TableBindings</span></span><br><span data-ttu-id="38f8d-169">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-169">
        - TableCoercion</span></span><br><span data-ttu-id="38f8d-170">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-170">
        - TextBindings</span></span><br><span data-ttu-id="38f8d-171">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-171">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-172">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="38f8d-172">Office 2019 on Windows</span></span><br><span data-ttu-id="38f8d-173">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="38f8d-173">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="38f8d-174">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="38f8d-174">- TaskPane</span></span><br><span data-ttu-id="38f8d-175">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="38f8d-175">
        - Content</span></span><br><span data-ttu-id="38f8d-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="38f8d-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="38f8d-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="38f8d-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="38f8d-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="38f8d-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="38f8d-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="38f8d-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="38f8d-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="38f8d-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="38f8d-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="38f8d-187">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-187">- BindingEvents</span></span><br><span data-ttu-id="38f8d-188">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-188">
        - CompressedFile</span></span><br><span data-ttu-id="38f8d-189">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-189">
        - DocumentEvents</span></span><br><span data-ttu-id="38f8d-190">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="38f8d-190">
        - File</span></span><br><span data-ttu-id="38f8d-191">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-191">
        - MatrixBindings</span></span><br><span data-ttu-id="38f8d-192">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-192">
        - MatrixCoercion</span></span><br><span data-ttu-id="38f8d-193">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="38f8d-193">
        - Selection</span></span><br><span data-ttu-id="38f8d-194">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="38f8d-194">
        - Settings</span></span><br><span data-ttu-id="38f8d-195">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-195">
        - TableBindings</span></span><br><span data-ttu-id="38f8d-196">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-196">
        - TableCoercion</span></span><br><span data-ttu-id="38f8d-197">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-197">
        - TextBindings</span></span><br><span data-ttu-id="38f8d-198">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-198">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-199">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="38f8d-199">Office 2016 on Windows</span></span><br><span data-ttu-id="38f8d-200">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="38f8d-200">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="38f8d-201">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="38f8d-201">- TaskPane</span></span><br><span data-ttu-id="38f8d-202">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="38f8d-202">
        - Content</span></span></td>
    <td><span data-ttu-id="38f8d-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="38f8d-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="38f8d-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="38f8d-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="38f8d-206">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-206">- BindingEvents</span></span><br><span data-ttu-id="38f8d-207">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-207">
        - CompressedFile</span></span><br><span data-ttu-id="38f8d-208">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-208">
        - DocumentEvents</span></span><br><span data-ttu-id="38f8d-209">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="38f8d-209">
        - File</span></span><br><span data-ttu-id="38f8d-210">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-210">
        - MatrixBindings</span></span><br><span data-ttu-id="38f8d-211">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-211">
        - MatrixCoercion</span></span><br><span data-ttu-id="38f8d-212">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="38f8d-212">
        - Selection</span></span><br><span data-ttu-id="38f8d-213">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="38f8d-213">
        - Settings</span></span><br><span data-ttu-id="38f8d-214">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-214">
        - TableBindings</span></span><br><span data-ttu-id="38f8d-215">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-215">
        - TableCoercion</span></span><br><span data-ttu-id="38f8d-216">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-216">
        - TextBindings</span></span><br><span data-ttu-id="38f8d-217">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-217">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-218">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="38f8d-218">Office 2013 on Windows</span></span><br><span data-ttu-id="38f8d-219">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="38f8d-219">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="38f8d-220">
        - TaskPane</span><span class="sxs-lookup"><span data-stu-id="38f8d-220">
        - TaskPane</span></span><br><span data-ttu-id="38f8d-221">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="38f8d-221">
        - Content</span></span></td>
    <td>  <span data-ttu-id="38f8d-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="38f8d-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="38f8d-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="38f8d-224">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-224">
        - BindingEvents</span></span><br><span data-ttu-id="38f8d-225">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-225">
        - CompressedFile</span></span><br><span data-ttu-id="38f8d-226">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-226">
        - DocumentEvents</span></span><br><span data-ttu-id="38f8d-227">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="38f8d-227">
        - File</span></span><br><span data-ttu-id="38f8d-228">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-228">
        - MatrixBindings</span></span><br><span data-ttu-id="38f8d-229">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-229">
        - MatrixCoercion</span></span><br><span data-ttu-id="38f8d-230">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="38f8d-230">
        - Selection</span></span><br><span data-ttu-id="38f8d-231">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="38f8d-231">
        - Settings</span></span><br><span data-ttu-id="38f8d-232">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-232">
        - TableBindings</span></span><br><span data-ttu-id="38f8d-233">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-233">
        - TableCoercion</span></span><br><span data-ttu-id="38f8d-234">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-234">
        - TextBindings</span></span><br><span data-ttu-id="38f8d-235">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-235">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-236">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="38f8d-236">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="38f8d-237">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="38f8d-237">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="38f8d-238">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="38f8d-238">- TaskPane</span></span><br><span data-ttu-id="38f8d-239">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="38f8d-239">
        - Content</span></span><br><span data-ttu-id="38f8d-240">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="38f8d-240">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="38f8d-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="38f8d-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="38f8d-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="38f8d-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="38f8d-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="38f8d-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="38f8d-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="38f8d-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="38f8d-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="38f8d-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="38f8d-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="38f8d-252">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-252">- BindingEvents</span></span><br><span data-ttu-id="38f8d-253">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-253">
        - DocumentEvents</span></span><br><span data-ttu-id="38f8d-254">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="38f8d-254">
        - File</span></span><br><span data-ttu-id="38f8d-255">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-255">
        - MatrixBindings</span></span><br><span data-ttu-id="38f8d-256">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-256">
        - MatrixCoercion</span></span><br><span data-ttu-id="38f8d-257">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="38f8d-257">
        - Selection</span></span><br><span data-ttu-id="38f8d-258">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="38f8d-258">
        - Settings</span></span><br><span data-ttu-id="38f8d-259">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-259">
        - TableBindings</span></span><br><span data-ttu-id="38f8d-260">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-260">
        - TableCoercion</span></span><br><span data-ttu-id="38f8d-261">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-261">
        - TextBindings</span></span><br><span data-ttu-id="38f8d-262">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-262">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-263">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="38f8d-263">Office apps on Mac</span></span><br><span data-ttu-id="38f8d-264">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="38f8d-264">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="38f8d-265">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="38f8d-265">- TaskPane</span></span><br><span data-ttu-id="38f8d-266">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="38f8d-266">
        - Content</span></span><br><span data-ttu-id="38f8d-267">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="38f8d-267">
        - Custom Functions</span></span><br><span data-ttu-id="38f8d-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="38f8d-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="38f8d-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="38f8d-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="38f8d-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="38f8d-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="38f8d-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="38f8d-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="38f8d-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="38f8d-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="38f8d-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="38f8d-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="38f8d-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="38f8d-281">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-281">- BindingEvents</span></span><br><span data-ttu-id="38f8d-282">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-282">
        - CompressedFile</span></span><br><span data-ttu-id="38f8d-283">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-283">
        - DocumentEvents</span></span><br><span data-ttu-id="38f8d-284">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="38f8d-284">
        - File</span></span><br><span data-ttu-id="38f8d-285">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-285">
        - MatrixBindings</span></span><br><span data-ttu-id="38f8d-286">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-286">
        - MatrixCoercion</span></span><br><span data-ttu-id="38f8d-287">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-287">
        - PdfFile</span></span><br><span data-ttu-id="38f8d-288">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="38f8d-288">
        - Selection</span></span><br><span data-ttu-id="38f8d-289">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="38f8d-289">
        - Settings</span></span><br><span data-ttu-id="38f8d-290">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-290">
        - TableBindings</span></span><br><span data-ttu-id="38f8d-291">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-291">
        - TableCoercion</span></span><br><span data-ttu-id="38f8d-292">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-292">
        - TextBindings</span></span><br><span data-ttu-id="38f8d-293">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-293">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-294">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="38f8d-294">Office 2019 for Mac</span></span><br><span data-ttu-id="38f8d-295">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="38f8d-295">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="38f8d-296">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="38f8d-296">- TaskPane</span></span><br><span data-ttu-id="38f8d-297">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="38f8d-297">
        - Content</span></span><br><span data-ttu-id="38f8d-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="38f8d-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="38f8d-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="38f8d-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="38f8d-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="38f8d-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="38f8d-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="38f8d-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="38f8d-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="38f8d-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="38f8d-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="38f8d-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-309">- BindingEvents</span></span><br><span data-ttu-id="38f8d-310">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-310">
        - CompressedFile</span></span><br><span data-ttu-id="38f8d-311">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-311">
        - DocumentEvents</span></span><br><span data-ttu-id="38f8d-312">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="38f8d-312">
        - File</span></span><br><span data-ttu-id="38f8d-313">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-313">
        - MatrixBindings</span></span><br><span data-ttu-id="38f8d-314">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-314">
        - MatrixCoercion</span></span><br><span data-ttu-id="38f8d-315">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-315">
        - PdfFile</span></span><br><span data-ttu-id="38f8d-316">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="38f8d-316">
        - Selection</span></span><br><span data-ttu-id="38f8d-317">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="38f8d-317">
        - Settings</span></span><br><span data-ttu-id="38f8d-318">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-318">
        - TableBindings</span></span><br><span data-ttu-id="38f8d-319">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-319">
        - TableCoercion</span></span><br><span data-ttu-id="38f8d-320">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-320">
        - TextBindings</span></span><br><span data-ttu-id="38f8d-321">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-321">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-322">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="38f8d-322">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="38f8d-323">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="38f8d-323">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="38f8d-324">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="38f8d-324">- TaskPane</span></span><br><span data-ttu-id="38f8d-325">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="38f8d-325">
        - Content</span></span></td>
    <td><span data-ttu-id="38f8d-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="38f8d-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="38f8d-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="38f8d-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="38f8d-329">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-329">- BindingEvents</span></span><br><span data-ttu-id="38f8d-330">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-330">
        - CompressedFile</span></span><br><span data-ttu-id="38f8d-331">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-331">
        - DocumentEvents</span></span><br><span data-ttu-id="38f8d-332">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="38f8d-332">
        - File</span></span><br><span data-ttu-id="38f8d-333">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-333">
        - MatrixBindings</span></span><br><span data-ttu-id="38f8d-334">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-334">
        - MatrixCoercion</span></span><br><span data-ttu-id="38f8d-335">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-335">
        - PdfFile</span></span><br><span data-ttu-id="38f8d-336">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="38f8d-336">
        - Selection</span></span><br><span data-ttu-id="38f8d-337">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="38f8d-337">
        - Settings</span></span><br><span data-ttu-id="38f8d-338">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-338">
        - TableBindings</span></span><br><span data-ttu-id="38f8d-339">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-339">
        - TableCoercion</span></span><br><span data-ttu-id="38f8d-340">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-340">
        - TextBindings</span></span><br><span data-ttu-id="38f8d-341">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-341">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="38f8d-342">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="38f8d-342">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="38f8d-343">Funções Personalizadas</span><span class="sxs-lookup"><span data-stu-id="38f8d-343">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="38f8d-344">Plataforma</span><span class="sxs-lookup"><span data-stu-id="38f8d-344">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="38f8d-345">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="38f8d-345">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="38f8d-346">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="38f8d-346">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="38f8d-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="38f8d-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-348">Office na Web</span><span class="sxs-lookup"><span data-stu-id="38f8d-348">Office on the web</span></span></td>
    <td><span data-ttu-id="38f8d-349">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="38f8d-349">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="38f8d-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-351">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="38f8d-351">Office on Windows</span></span><br><span data-ttu-id="38f8d-352">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="38f8d-352">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="38f8d-353">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="38f8d-353">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="38f8d-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-355">Office para Mac</span><span class="sxs-lookup"><span data-stu-id="38f8d-355">Office for Mac</span></span><br><span data-ttu-id="38f8d-356">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="38f8d-356">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="38f8d-357">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="38f8d-357">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="38f8d-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="38f8d-359">Outlook</span><span class="sxs-lookup"><span data-stu-id="38f8d-359">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="38f8d-360">Plataforma</span><span class="sxs-lookup"><span data-stu-id="38f8d-360">Platform</span></span></th>
    <th><span data-ttu-id="38f8d-361">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="38f8d-361">Extension points</span></span></th>
    <th><span data-ttu-id="38f8d-362">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="38f8d-362">API requirement sets</span></span></th>
    <th><span data-ttu-id="38f8d-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="38f8d-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-364">Office na Web</span><span class="sxs-lookup"><span data-stu-id="38f8d-364">Office on the web</span></span><br><span data-ttu-id="38f8d-365">(novo)</span><span class="sxs-lookup"><span data-stu-id="38f8d-365">New</span></span></td>
    <td> <span data-ttu-id="38f8d-366">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="38f8d-366">- Mail Read</span></span><br><span data-ttu-id="38f8d-367">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="38f8d-367">
      - Mail Compose</span></span><br><span data-ttu-id="38f8d-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="38f8d-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="38f8d-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="38f8d-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="38f8d-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="38f8d-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="38f8d-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="38f8d-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="38f8d-376">Não disponível</span><span class="sxs-lookup"><span data-stu-id="38f8d-376">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-377">Office na Web</span><span class="sxs-lookup"><span data-stu-id="38f8d-377">Office on the web</span></span><br><span data-ttu-id="38f8d-378">(clássico)</span><span class="sxs-lookup"><span data-stu-id="38f8d-378">Classic.</span></span></td>
    <td> <span data-ttu-id="38f8d-379">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="38f8d-379">- Mail Read</span></span><br><span data-ttu-id="38f8d-380">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="38f8d-380">
      - Mail Compose</span></span><br><span data-ttu-id="38f8d-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="38f8d-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="38f8d-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="38f8d-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="38f8d-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="38f8d-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="38f8d-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="38f8d-388">Não disponível</span><span class="sxs-lookup"><span data-stu-id="38f8d-388">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-389">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="38f8d-389">Office on Windows</span></span><br><span data-ttu-id="38f8d-390">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="38f8d-390">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="38f8d-391">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="38f8d-391">- Mail Read</span></span><br><span data-ttu-id="38f8d-392">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="38f8d-392">
      - Mail Compose</span></span><br><span data-ttu-id="38f8d-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="38f8d-394">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="38f8d-394">
      - Modules</span></span></td>
    <td> <span data-ttu-id="38f8d-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="38f8d-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="38f8d-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="38f8d-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="38f8d-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="38f8d-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="38f8d-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="38f8d-402">Não disponível</span><span class="sxs-lookup"><span data-stu-id="38f8d-402">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-403">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="38f8d-403">Office 2019 on Windows</span></span><br><span data-ttu-id="38f8d-404">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="38f8d-404">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="38f8d-405">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="38f8d-405">- Mail Read</span></span><br><span data-ttu-id="38f8d-406">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="38f8d-406">
      - Mail Compose</span></span><br><span data-ttu-id="38f8d-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="38f8d-408">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="38f8d-408">
      - Modules</span></span></td>
    <td> <span data-ttu-id="38f8d-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="38f8d-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="38f8d-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="38f8d-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="38f8d-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="38f8d-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="38f8d-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="38f8d-416">Não disponível</span><span class="sxs-lookup"><span data-stu-id="38f8d-416">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-417">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="38f8d-417">Office 2016 on Windows</span></span><br><span data-ttu-id="38f8d-418">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="38f8d-418">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="38f8d-419">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="38f8d-419">- Mail Read</span></span><br><span data-ttu-id="38f8d-420">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="38f8d-420">
      - Mail Compose</span></span><br><span data-ttu-id="38f8d-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="38f8d-422">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="38f8d-422">
      - Modules</span></span></td>
    <td> <span data-ttu-id="38f8d-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="38f8d-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="38f8d-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="38f8d-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="38f8d-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="38f8d-427">Não disponível</span><span class="sxs-lookup"><span data-stu-id="38f8d-427">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-428">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="38f8d-428">Office 2013 on Windows</span></span><br><span data-ttu-id="38f8d-429">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="38f8d-429">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="38f8d-430">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="38f8d-430">- Mail Read</span></span><br><span data-ttu-id="38f8d-431">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="38f8d-431">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="38f8d-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="38f8d-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="38f8d-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="38f8d-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="38f8d-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="38f8d-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="38f8d-436">Não disponível</span><span class="sxs-lookup"><span data-stu-id="38f8d-436">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-437">Office no iOS</span><span class="sxs-lookup"><span data-stu-id="38f8d-437">Office apps on iOS</span></span><br><span data-ttu-id="38f8d-438">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="38f8d-438">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="38f8d-439">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="38f8d-439">- Mail Read</span></span><br><span data-ttu-id="38f8d-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="38f8d-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="38f8d-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="38f8d-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="38f8d-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="38f8d-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="38f8d-446">Não disponível</span><span class="sxs-lookup"><span data-stu-id="38f8d-446">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-447">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="38f8d-447">Office apps on Mac</span></span><br><span data-ttu-id="38f8d-448">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="38f8d-448">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="38f8d-449">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="38f8d-449">- Mail Read</span></span><br><span data-ttu-id="38f8d-450">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="38f8d-450">
      - Mail Compose</span></span><br><span data-ttu-id="38f8d-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="38f8d-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="38f8d-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="38f8d-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="38f8d-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="38f8d-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="38f8d-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="38f8d-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="38f8d-459">Não disponível</span><span class="sxs-lookup"><span data-stu-id="38f8d-459">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-460">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="38f8d-460">Office 2019 for Mac</span></span><br><span data-ttu-id="38f8d-461">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="38f8d-461">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="38f8d-462">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="38f8d-462">- Mail Read</span></span><br><span data-ttu-id="38f8d-463">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="38f8d-463">
      - Mail Compose</span></span><br><span data-ttu-id="38f8d-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="38f8d-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="38f8d-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="38f8d-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="38f8d-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="38f8d-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="38f8d-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="38f8d-471">Não disponível</span><span class="sxs-lookup"><span data-stu-id="38f8d-471">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-472">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="38f8d-472">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="38f8d-473">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="38f8d-473">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="38f8d-474">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="38f8d-474">- Mail Read</span></span><br><span data-ttu-id="38f8d-475">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="38f8d-475">
      - Mail Compose</span></span><br><span data-ttu-id="38f8d-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="38f8d-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="38f8d-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="38f8d-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="38f8d-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="38f8d-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="38f8d-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="38f8d-483">Não disponível</span><span class="sxs-lookup"><span data-stu-id="38f8d-483">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-484">Outlook no Android</span><span class="sxs-lookup"><span data-stu-id="38f8d-484">Office apps on Android</span></span><br><span data-ttu-id="38f8d-485">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="38f8d-485">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="38f8d-486">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="38f8d-486">- Mail Read</span></span><br><span data-ttu-id="38f8d-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="38f8d-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="38f8d-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="38f8d-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="38f8d-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="38f8d-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="38f8d-493">Não disponível</span><span class="sxs-lookup"><span data-stu-id="38f8d-493">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="38f8d-494">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="38f8d-494">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="38f8d-495">Word</span><span class="sxs-lookup"><span data-stu-id="38f8d-495">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="38f8d-496">Plataforma</span><span class="sxs-lookup"><span data-stu-id="38f8d-496">Platform</span></span></th>
    <th><span data-ttu-id="38f8d-497">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="38f8d-497">Extension points</span></span></th>
    <th><span data-ttu-id="38f8d-498">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="38f8d-498">API requirement sets</span></span></th>
    <th><span data-ttu-id="38f8d-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="38f8d-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-500">Office na Web</span><span class="sxs-lookup"><span data-stu-id="38f8d-500">Office on the web</span></span></td>
    <td> <span data-ttu-id="38f8d-501">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="38f8d-501">- TaskPane</span></span><br><span data-ttu-id="38f8d-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="38f8d-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="38f8d-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="38f8d-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="38f8d-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="38f8d-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="38f8d-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="38f8d-509">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-509">- BindingEvents</span></span><br><span data-ttu-id="38f8d-510">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="38f8d-510">
         - CustomXmlParts</span></span><br><span data-ttu-id="38f8d-511">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-511">
         - DocumentEvents</span></span><br><span data-ttu-id="38f8d-512">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="38f8d-512">
         - File</span></span><br><span data-ttu-id="38f8d-513">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-513">
         - HtmlCoercion</span></span><br><span data-ttu-id="38f8d-514">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-514">
         - MatrixBindings</span></span><br><span data-ttu-id="38f8d-515">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-515">
         - MatrixCoercion</span></span><br><span data-ttu-id="38f8d-516">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-516">
         - OoxmlCoercion</span></span><br><span data-ttu-id="38f8d-517">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-517">
         - PdfFile</span></span><br><span data-ttu-id="38f8d-518">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="38f8d-518">
         - Selection</span></span><br><span data-ttu-id="38f8d-519">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="38f8d-519">
         - Settings</span></span><br><span data-ttu-id="38f8d-520">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-520">
         - TableBindings</span></span><br><span data-ttu-id="38f8d-521">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-521">
         - TableCoercion</span></span><br><span data-ttu-id="38f8d-522">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-522">
         - TextBindings</span></span><br><span data-ttu-id="38f8d-523">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-523">
         - TextCoercion</span></span><br><span data-ttu-id="38f8d-524">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-524">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-525">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="38f8d-525">Office on Windows</span></span><br><span data-ttu-id="38f8d-526">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="38f8d-526">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="38f8d-527">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="38f8d-527">- TaskPane</span></span><br><span data-ttu-id="38f8d-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="38f8d-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="38f8d-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="38f8d-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="38f8d-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="38f8d-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="38f8d-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="38f8d-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-535">- BindingEvents</span></span><br><span data-ttu-id="38f8d-536">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-536">
         - CompressedFile</span></span><br><span data-ttu-id="38f8d-537">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="38f8d-537">
         - CustomXmlParts</span></span><br><span data-ttu-id="38f8d-538">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-538">
         - DocumentEvents</span></span><br><span data-ttu-id="38f8d-539">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="38f8d-539">
         - File</span></span><br><span data-ttu-id="38f8d-540">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-540">
         - HtmlCoercion</span></span><br><span data-ttu-id="38f8d-541">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-541">
         - MatrixBindings</span></span><br><span data-ttu-id="38f8d-542">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-542">
         - MatrixCoercion</span></span><br><span data-ttu-id="38f8d-543">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-543">
         - OoxmlCoercion</span></span><br><span data-ttu-id="38f8d-544">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-544">
         - PdfFile</span></span><br><span data-ttu-id="38f8d-545">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="38f8d-545">
         - Selection</span></span><br><span data-ttu-id="38f8d-546">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="38f8d-546">
         - Settings</span></span><br><span data-ttu-id="38f8d-547">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-547">
         - TableBindings</span></span><br><span data-ttu-id="38f8d-548">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-548">
         - TableCoercion</span></span><br><span data-ttu-id="38f8d-549">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-549">
         - TextBindings</span></span><br><span data-ttu-id="38f8d-550">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-550">
         - TextCoercion</span></span><br><span data-ttu-id="38f8d-551">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-551">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-552">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="38f8d-552">Office 2019 on Windows</span></span><br><span data-ttu-id="38f8d-553">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="38f8d-553">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="38f8d-554">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="38f8d-554">- TaskPane</span></span><br><span data-ttu-id="38f8d-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="38f8d-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="38f8d-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="38f8d-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="38f8d-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="38f8d-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="38f8d-561">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-561">- BindingEvents</span></span><br><span data-ttu-id="38f8d-562">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-562">
         - CompressedFile</span></span><br><span data-ttu-id="38f8d-563">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="38f8d-563">
         - CustomXmlParts</span></span><br><span data-ttu-id="38f8d-564">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-564">
         - DocumentEvents</span></span><br><span data-ttu-id="38f8d-565">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="38f8d-565">
         - File</span></span><br><span data-ttu-id="38f8d-566">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-566">
         - HtmlCoercion</span></span><br><span data-ttu-id="38f8d-567">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-567">
         - MatrixBindings</span></span><br><span data-ttu-id="38f8d-568">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-568">
         - MatrixCoercion</span></span><br><span data-ttu-id="38f8d-569">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-569">
         - OoxmlCoercion</span></span><br><span data-ttu-id="38f8d-570">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-570">
         - PdfFile</span></span><br><span data-ttu-id="38f8d-571">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="38f8d-571">
         - Selection</span></span><br><span data-ttu-id="38f8d-572">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="38f8d-572">
         - Settings</span></span><br><span data-ttu-id="38f8d-573">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-573">
         - TableBindings</span></span><br><span data-ttu-id="38f8d-574">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-574">
         - TableCoercion</span></span><br><span data-ttu-id="38f8d-575">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-575">
         - TextBindings</span></span><br><span data-ttu-id="38f8d-576">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-576">
         - TextCoercion</span></span><br><span data-ttu-id="38f8d-577">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-577">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-578">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="38f8d-578">Office 2016 on Windows</span></span><br><span data-ttu-id="38f8d-579">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="38f8d-579">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="38f8d-580">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="38f8d-580">- TaskPane</span></span></td>
    <td> <span data-ttu-id="38f8d-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="38f8d-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="38f8d-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="38f8d-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="38f8d-584">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-584">- BindingEvents</span></span><br><span data-ttu-id="38f8d-585">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-585">
         - CompressedFile</span></span><br><span data-ttu-id="38f8d-586">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="38f8d-586">
         - CustomXmlParts</span></span><br><span data-ttu-id="38f8d-587">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-587">
         - DocumentEvents</span></span><br><span data-ttu-id="38f8d-588">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="38f8d-588">
         - File</span></span><br><span data-ttu-id="38f8d-589">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-589">
         - HtmlCoercion</span></span><br><span data-ttu-id="38f8d-590">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-590">
         - MatrixBindings</span></span><br><span data-ttu-id="38f8d-591">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-591">
         - MatrixCoercion</span></span><br><span data-ttu-id="38f8d-592">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-592">
         - OoxmlCoercion</span></span><br><span data-ttu-id="38f8d-593">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-593">
         - PdfFile</span></span><br><span data-ttu-id="38f8d-594">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="38f8d-594">
         - Selection</span></span><br><span data-ttu-id="38f8d-595">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="38f8d-595">
         - Settings</span></span><br><span data-ttu-id="38f8d-596">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-596">
         - TableBindings</span></span><br><span data-ttu-id="38f8d-597">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-597">
         - TableCoercion</span></span><br><span data-ttu-id="38f8d-598">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-598">
         - TextBindings</span></span><br><span data-ttu-id="38f8d-599">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-599">
         - TextCoercion</span></span><br><span data-ttu-id="38f8d-600">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-600">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-601">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="38f8d-601">Office 2013 on Windows</span></span><br><span data-ttu-id="38f8d-602">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="38f8d-602">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="38f8d-603">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="38f8d-603">- TaskPane</span></span></td>
    <td> <span data-ttu-id="38f8d-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="38f8d-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="38f8d-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="38f8d-606">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-606">- BindingEvents</span></span><br><span data-ttu-id="38f8d-607">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-607">
         - CompressedFile</span></span><br><span data-ttu-id="38f8d-608">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="38f8d-608">
         - CustomXmlParts</span></span><br><span data-ttu-id="38f8d-609">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-609">
         - DocumentEvents</span></span><br><span data-ttu-id="38f8d-610">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="38f8d-610">
         - File</span></span><br><span data-ttu-id="38f8d-611">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-611">
         - HtmlCoercion</span></span><br><span data-ttu-id="38f8d-612">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-612">
         - MatrixBindings</span></span><br><span data-ttu-id="38f8d-613">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-613">
         - MatrixCoercion</span></span><br><span data-ttu-id="38f8d-614">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-614">
         - OoxmlCoercion</span></span><br><span data-ttu-id="38f8d-615">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-615">
         - PdfFile</span></span><br><span data-ttu-id="38f8d-616">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="38f8d-616">
         - Selection</span></span><br><span data-ttu-id="38f8d-617">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="38f8d-617">
         - Settings</span></span><br><span data-ttu-id="38f8d-618">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-618">
         - TableBindings</span></span><br><span data-ttu-id="38f8d-619">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-619">
         - TableCoercion</span></span><br><span data-ttu-id="38f8d-620">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-620">
         - TextBindings</span></span><br><span data-ttu-id="38f8d-621">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-621">
         - TextCoercion</span></span><br><span data-ttu-id="38f8d-622">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-622">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-623">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="38f8d-623">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="38f8d-624">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="38f8d-624">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="38f8d-625">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="38f8d-625">- TaskPane</span></span></td>
    <td> <span data-ttu-id="38f8d-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="38f8d-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="38f8d-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="38f8d-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="38f8d-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="38f8d-631">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-631">- BindingEvents</span></span><br><span data-ttu-id="38f8d-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-632">
         - CompressedFile</span></span><br><span data-ttu-id="38f8d-633">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="38f8d-633">
         - CustomXmlParts</span></span><br><span data-ttu-id="38f8d-634">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-634">
         - DocumentEvents</span></span><br><span data-ttu-id="38f8d-635">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="38f8d-635">
         - File</span></span><br><span data-ttu-id="38f8d-636">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-636">
         - HtmlCoercion</span></span><br><span data-ttu-id="38f8d-637">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-637">
         - MatrixBindings</span></span><br><span data-ttu-id="38f8d-638">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-638">
         - MatrixCoercion</span></span><br><span data-ttu-id="38f8d-639">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-639">
         - OoxmlCoercion</span></span><br><span data-ttu-id="38f8d-640">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-640">
         - PdfFile</span></span><br><span data-ttu-id="38f8d-641">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="38f8d-641">
         - Selection</span></span><br><span data-ttu-id="38f8d-642">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="38f8d-642">
         - Settings</span></span><br><span data-ttu-id="38f8d-643">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-643">
         - TableBindings</span></span><br><span data-ttu-id="38f8d-644">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-644">
         - TableCoercion</span></span><br><span data-ttu-id="38f8d-645">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-645">
         - TextBindings</span></span><br><span data-ttu-id="38f8d-646">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-646">
         - TextCoercion</span></span><br><span data-ttu-id="38f8d-647">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-647">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-648">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="38f8d-648">Office apps on Mac</span></span><br><span data-ttu-id="38f8d-649">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="38f8d-649">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="38f8d-650">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="38f8d-650">- TaskPane</span></span><br><span data-ttu-id="38f8d-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="38f8d-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="38f8d-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="38f8d-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="38f8d-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="38f8d-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="38f8d-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="38f8d-658">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-658">- BindingEvents</span></span><br><span data-ttu-id="38f8d-659">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-659">
         - CompressedFile</span></span><br><span data-ttu-id="38f8d-660">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="38f8d-660">
         - CustomXmlParts</span></span><br><span data-ttu-id="38f8d-661">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-661">
         - DocumentEvents</span></span><br><span data-ttu-id="38f8d-662">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="38f8d-662">
         - File</span></span><br><span data-ttu-id="38f8d-663">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-663">
         - HtmlCoercion</span></span><br><span data-ttu-id="38f8d-664">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-664">
         - MatrixBindings</span></span><br><span data-ttu-id="38f8d-665">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-665">
         - MatrixCoercion</span></span><br><span data-ttu-id="38f8d-666">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-666">
         - OoxmlCoercion</span></span><br><span data-ttu-id="38f8d-667">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-667">
         - PdfFile</span></span><br><span data-ttu-id="38f8d-668">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="38f8d-668">
         - Selection</span></span><br><span data-ttu-id="38f8d-669">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="38f8d-669">
         - Settings</span></span><br><span data-ttu-id="38f8d-670">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-670">
         - TableBindings</span></span><br><span data-ttu-id="38f8d-671">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-671">
         - TableCoercion</span></span><br><span data-ttu-id="38f8d-672">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-672">
         - TextBindings</span></span><br><span data-ttu-id="38f8d-673">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-673">
         - TextCoercion</span></span><br><span data-ttu-id="38f8d-674">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-674">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-675">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="38f8d-675">Office 2019 for Mac</span></span><br><span data-ttu-id="38f8d-676">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="38f8d-676">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="38f8d-677">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="38f8d-677">- TaskPane</span></span><br><span data-ttu-id="38f8d-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="38f8d-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="38f8d-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="38f8d-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="38f8d-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="38f8d-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="38f8d-684">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-684">- BindingEvents</span></span><br><span data-ttu-id="38f8d-685">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-685">
         - CompressedFile</span></span><br><span data-ttu-id="38f8d-686">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="38f8d-686">
         - CustomXmlParts</span></span><br><span data-ttu-id="38f8d-687">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-687">
         - DocumentEvents</span></span><br><span data-ttu-id="38f8d-688">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="38f8d-688">
         - File</span></span><br><span data-ttu-id="38f8d-689">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-689">
         - HtmlCoercion</span></span><br><span data-ttu-id="38f8d-690">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-690">
         - MatrixBindings</span></span><br><span data-ttu-id="38f8d-691">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-691">
         - MatrixCoercion</span></span><br><span data-ttu-id="38f8d-692">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-692">
         - OoxmlCoercion</span></span><br><span data-ttu-id="38f8d-693">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-693">
         - PdfFile</span></span><br><span data-ttu-id="38f8d-694">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="38f8d-694">
         - Selection</span></span><br><span data-ttu-id="38f8d-695">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="38f8d-695">
         - Settings</span></span><br><span data-ttu-id="38f8d-696">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-696">
         - TableBindings</span></span><br><span data-ttu-id="38f8d-697">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-697">
         - TableCoercion</span></span><br><span data-ttu-id="38f8d-698">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-698">
         - TextBindings</span></span><br><span data-ttu-id="38f8d-699">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-699">
         - TextCoercion</span></span><br><span data-ttu-id="38f8d-700">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-700">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-701">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="38f8d-701">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="38f8d-702">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="38f8d-702">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="38f8d-703">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="38f8d-703">- TaskPane</span></span></td>
    <td> <span data-ttu-id="38f8d-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="38f8d-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="38f8d-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="38f8d-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="38f8d-707">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-707">- BindingEvents</span></span><br><span data-ttu-id="38f8d-708">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-708">
         - CompressedFile</span></span><br><span data-ttu-id="38f8d-709">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="38f8d-709">
         - CustomXmlParts</span></span><br><span data-ttu-id="38f8d-710">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-710">
         - DocumentEvents</span></span><br><span data-ttu-id="38f8d-711">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="38f8d-711">
         - File</span></span><br><span data-ttu-id="38f8d-712">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-712">
         - HtmlCoercion</span></span><br><span data-ttu-id="38f8d-713">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-713">
         - MatrixBindings</span></span><br><span data-ttu-id="38f8d-714">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-714">
         - MatrixCoercion</span></span><br><span data-ttu-id="38f8d-715">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-715">
         - OoxmlCoercion</span></span><br><span data-ttu-id="38f8d-716">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-716">
         - PdfFile</span></span><br><span data-ttu-id="38f8d-717">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="38f8d-717">
         - Selection</span></span><br><span data-ttu-id="38f8d-718">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="38f8d-718">
         - Settings</span></span><br><span data-ttu-id="38f8d-719">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-719">
         - TableBindings</span></span><br><span data-ttu-id="38f8d-720">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-720">
         - TableCoercion</span></span><br><span data-ttu-id="38f8d-721">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="38f8d-721">
         - TextBindings</span></span><br><span data-ttu-id="38f8d-722">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-722">
         - TextCoercion</span></span><br><span data-ttu-id="38f8d-723">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-723">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="38f8d-724">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="38f8d-724">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="38f8d-725">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="38f8d-725">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="38f8d-726">Plataforma</span><span class="sxs-lookup"><span data-stu-id="38f8d-726">Platform</span></span></th>
    <th><span data-ttu-id="38f8d-727">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="38f8d-727">Extension points</span></span></th>
    <th><span data-ttu-id="38f8d-728">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="38f8d-728">API requirement sets</span></span></th>
    <th><span data-ttu-id="38f8d-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="38f8d-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-730">Office na Web</span><span class="sxs-lookup"><span data-stu-id="38f8d-730">Office on the web</span></span></td>
    <td> <span data-ttu-id="38f8d-731">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="38f8d-731">- Content</span></span><br><span data-ttu-id="38f8d-732">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="38f8d-732">
         - TaskPane</span></span><br><span data-ttu-id="38f8d-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="38f8d-734">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-734">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="38f8d-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="38f8d-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="38f8d-737">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="38f8d-737">- ActiveView</span></span><br><span data-ttu-id="38f8d-738">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-738">
         - CompressedFile</span></span><br><span data-ttu-id="38f8d-739">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-739">
         - DocumentEvents</span></span><br><span data-ttu-id="38f8d-740">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="38f8d-740">
         - File</span></span><br><span data-ttu-id="38f8d-741">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-741">
         - PdfFile</span></span><br><span data-ttu-id="38f8d-742">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="38f8d-742">
         - Selection</span></span><br><span data-ttu-id="38f8d-743">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="38f8d-743">
         - Settings</span></span><br><span data-ttu-id="38f8d-744">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-744">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-745">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="38f8d-745">Office on Windows</span></span><br><span data-ttu-id="38f8d-746">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="38f8d-746">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="38f8d-747">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="38f8d-747">- Content</span></span><br><span data-ttu-id="38f8d-748">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="38f8d-748">
         - TaskPane</span></span><br><span data-ttu-id="38f8d-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="38f8d-750">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-750">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="38f8d-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="38f8d-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="38f8d-753">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="38f8d-753">- ActiveView</span></span><br><span data-ttu-id="38f8d-754">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-754">
         - CompressedFile</span></span><br><span data-ttu-id="38f8d-755">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-755">
         - DocumentEvents</span></span><br><span data-ttu-id="38f8d-756">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="38f8d-756">
         - File</span></span><br><span data-ttu-id="38f8d-757">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-757">
         - PdfFile</span></span><br><span data-ttu-id="38f8d-758">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="38f8d-758">
         - Selection</span></span><br><span data-ttu-id="38f8d-759">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="38f8d-759">
         - Settings</span></span><br><span data-ttu-id="38f8d-760">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-760">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-761">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="38f8d-761">Office 2019 on Windows</span></span><br><span data-ttu-id="38f8d-762">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="38f8d-762">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="38f8d-763">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="38f8d-763">- Content</span></span><br><span data-ttu-id="38f8d-764">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="38f8d-764">
         - TaskPane</span></span><br><span data-ttu-id="38f8d-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="38f8d-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="38f8d-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="38f8d-768">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="38f8d-768">- ActiveView</span></span><br><span data-ttu-id="38f8d-769">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-769">
         - CompressedFile</span></span><br><span data-ttu-id="38f8d-770">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-770">
         - DocumentEvents</span></span><br><span data-ttu-id="38f8d-771">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="38f8d-771">
         - File</span></span><br><span data-ttu-id="38f8d-772">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-772">
         - PdfFile</span></span><br><span data-ttu-id="38f8d-773">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="38f8d-773">
         - Selection</span></span><br><span data-ttu-id="38f8d-774">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="38f8d-774">
         - Settings</span></span><br><span data-ttu-id="38f8d-775">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-775">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-776">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="38f8d-776">Office 2016 on Windows</span></span><br><span data-ttu-id="38f8d-777">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="38f8d-777">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="38f8d-778">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="38f8d-778">- Content</span></span><br><span data-ttu-id="38f8d-779">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="38f8d-779">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="38f8d-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="38f8d-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="38f8d-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="38f8d-782">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="38f8d-782">- ActiveView</span></span><br><span data-ttu-id="38f8d-783">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-783">
         - CompressedFile</span></span><br><span data-ttu-id="38f8d-784">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-784">
         - DocumentEvents</span></span><br><span data-ttu-id="38f8d-785">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="38f8d-785">
         - File</span></span><br><span data-ttu-id="38f8d-786">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-786">
         - PdfFile</span></span><br><span data-ttu-id="38f8d-787">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="38f8d-787">
         - Selection</span></span><br><span data-ttu-id="38f8d-788">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="38f8d-788">
         - Settings</span></span><br><span data-ttu-id="38f8d-789">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-789">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-790">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="38f8d-790">Office 2013 on Windows</span></span><br><span data-ttu-id="38f8d-791">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="38f8d-791">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="38f8d-792">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="38f8d-792">- Content</span></span><br><span data-ttu-id="38f8d-793">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="38f8d-793">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="38f8d-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="38f8d-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="38f8d-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="38f8d-796">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="38f8d-796">- ActiveView</span></span><br><span data-ttu-id="38f8d-797">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-797">
         - CompressedFile</span></span><br><span data-ttu-id="38f8d-798">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-798">
         - DocumentEvents</span></span><br><span data-ttu-id="38f8d-799">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="38f8d-799">
         - File</span></span><br><span data-ttu-id="38f8d-800">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-800">
         - PdfFile</span></span><br><span data-ttu-id="38f8d-801">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="38f8d-801">
         - Selection</span></span><br><span data-ttu-id="38f8d-802">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="38f8d-802">
         - Settings</span></span><br><span data-ttu-id="38f8d-803">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-803">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-804">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="38f8d-804">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="38f8d-805">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="38f8d-805">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="38f8d-806">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="38f8d-806">- Content</span></span><br><span data-ttu-id="38f8d-807">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="38f8d-807">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="38f8d-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="38f8d-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="38f8d-810">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="38f8d-810">- ActiveView</span></span><br><span data-ttu-id="38f8d-811">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-811">
         - CompressedFile</span></span><br><span data-ttu-id="38f8d-812">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-812">
         - DocumentEvents</span></span><br><span data-ttu-id="38f8d-813">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="38f8d-813">
         - File</span></span><br><span data-ttu-id="38f8d-814">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-814">
         - PdfFile</span></span><br><span data-ttu-id="38f8d-815">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="38f8d-815">
         - Selection</span></span><br><span data-ttu-id="38f8d-816">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="38f8d-816">
         - Settings</span></span><br><span data-ttu-id="38f8d-817">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-817">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-818">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="38f8d-818">Office apps on Mac</span></span><br><span data-ttu-id="38f8d-819">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="38f8d-819">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="38f8d-820">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="38f8d-820">- Content</span></span><br><span data-ttu-id="38f8d-821">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="38f8d-821">
         - TaskPane</span></span><br><span data-ttu-id="38f8d-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="38f8d-823">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-823">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="38f8d-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="38f8d-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="38f8d-826">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="38f8d-826">- ActiveView</span></span><br><span data-ttu-id="38f8d-827">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-827">
         - CompressedFile</span></span><br><span data-ttu-id="38f8d-828">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-828">
         - DocumentEvents</span></span><br><span data-ttu-id="38f8d-829">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="38f8d-829">
         - File</span></span><br><span data-ttu-id="38f8d-830">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-830">
         - PdfFile</span></span><br><span data-ttu-id="38f8d-831">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="38f8d-831">
         - Selection</span></span><br><span data-ttu-id="38f8d-832">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="38f8d-832">
         - Settings</span></span><br><span data-ttu-id="38f8d-833">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-833">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-834">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="38f8d-834">Office 2019 for Mac</span></span><br><span data-ttu-id="38f8d-835">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="38f8d-835">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="38f8d-836">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="38f8d-836">- Content</span></span><br><span data-ttu-id="38f8d-837">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="38f8d-837">
         - TaskPane</span></span><br><span data-ttu-id="38f8d-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="38f8d-839">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-839">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="38f8d-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="38f8d-841">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="38f8d-841">- ActiveView</span></span><br><span data-ttu-id="38f8d-842">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-842">
         - CompressedFile</span></span><br><span data-ttu-id="38f8d-843">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-843">
         - DocumentEvents</span></span><br><span data-ttu-id="38f8d-844">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="38f8d-844">
         - File</span></span><br><span data-ttu-id="38f8d-845">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-845">
         - PdfFile</span></span><br><span data-ttu-id="38f8d-846">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="38f8d-846">
         - Selection</span></span><br><span data-ttu-id="38f8d-847">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="38f8d-847">
         - Settings</span></span><br><span data-ttu-id="38f8d-848">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-848">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-849">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="38f8d-849">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="38f8d-850">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="38f8d-850">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="38f8d-851">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="38f8d-851">- Content</span></span><br><span data-ttu-id="38f8d-852">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="38f8d-852">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="38f8d-853">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="38f8d-853">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="38f8d-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="38f8d-855">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="38f8d-855">- ActiveView</span></span><br><span data-ttu-id="38f8d-856">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-856">
         - CompressedFile</span></span><br><span data-ttu-id="38f8d-857">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-857">
         - DocumentEvents</span></span><br><span data-ttu-id="38f8d-858">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="38f8d-858">
         - File</span></span><br><span data-ttu-id="38f8d-859">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="38f8d-859">
         - PdfFile</span></span><br><span data-ttu-id="38f8d-860">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="38f8d-860">
         - Selection</span></span><br><span data-ttu-id="38f8d-861">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="38f8d-861">
         - Settings</span></span><br><span data-ttu-id="38f8d-862">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-862">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="38f8d-863">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="38f8d-863">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="38f8d-864">OneNote</span><span class="sxs-lookup"><span data-stu-id="38f8d-864">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="38f8d-865">Plataforma</span><span class="sxs-lookup"><span data-stu-id="38f8d-865">Platform</span></span></th>
    <th><span data-ttu-id="38f8d-866">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="38f8d-866">Extension points</span></span></th>
    <th><span data-ttu-id="38f8d-867">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="38f8d-867">API requirement sets</span></span></th>
    <th><span data-ttu-id="38f8d-868"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="38f8d-868"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-869">Office na Web</span><span class="sxs-lookup"><span data-stu-id="38f8d-869">Office on the web</span></span></td>
    <td> <span data-ttu-id="38f8d-870">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="38f8d-870">- Content</span></span><br><span data-ttu-id="38f8d-871">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="38f8d-871">
         - TaskPane</span></span><br><span data-ttu-id="38f8d-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="38f8d-873">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-873">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="38f8d-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="38f8d-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="38f8d-876">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="38f8d-876">- DocumentEvents</span></span><br><span data-ttu-id="38f8d-877">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-877">
         - HtmlCoercion</span></span><br><span data-ttu-id="38f8d-878">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="38f8d-878">
         - Settings</span></span><br><span data-ttu-id="38f8d-879">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-879">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="38f8d-880">Project</span><span class="sxs-lookup"><span data-stu-id="38f8d-880">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="38f8d-881">Plataforma</span><span class="sxs-lookup"><span data-stu-id="38f8d-881">Platform</span></span></th>
    <th><span data-ttu-id="38f8d-882">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="38f8d-882">Extension points</span></span></th>
    <th><span data-ttu-id="38f8d-883">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="38f8d-883">API requirement sets</span></span></th>
    <th><span data-ttu-id="38f8d-884"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="38f8d-884"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-885">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="38f8d-885">Office 2019 on Windows</span></span><br><span data-ttu-id="38f8d-886">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="38f8d-886">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="38f8d-887">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="38f8d-887">- TaskPane</span></span></td>
    <td> <span data-ttu-id="38f8d-888">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-888">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="38f8d-889">- Seleção</span><span class="sxs-lookup"><span data-stu-id="38f8d-889">- Selection</span></span><br><span data-ttu-id="38f8d-890">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-890">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-891">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="38f8d-891">Office 2016 on Windows</span></span><br><span data-ttu-id="38f8d-892">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="38f8d-892">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="38f8d-893">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="38f8d-893">- TaskPane</span></span></td>
    <td> <span data-ttu-id="38f8d-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="38f8d-895">- Seleção</span><span class="sxs-lookup"><span data-stu-id="38f8d-895">- Selection</span></span><br><span data-ttu-id="38f8d-896">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-896">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="38f8d-897">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="38f8d-897">Office 2013 on Windows</span></span><br><span data-ttu-id="38f8d-898">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="38f8d-898">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="38f8d-899">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="38f8d-899">- TaskPane</span></span></td>
    <td> <span data-ttu-id="38f8d-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="38f8d-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="38f8d-901">- Seleção</span><span class="sxs-lookup"><span data-stu-id="38f8d-901">- Selection</span></span><br><span data-ttu-id="38f8d-902">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="38f8d-902">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="38f8d-903">Confira também</span><span class="sxs-lookup"><span data-stu-id="38f8d-903">See also</span></span>

- [<span data-ttu-id="38f8d-904">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="38f8d-904">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="38f8d-905">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="38f8d-905">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="38f8d-906">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="38f8d-906">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="38f8d-907">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="38f8d-907">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="38f8d-908">Referência da API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="38f8d-908">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="38f8d-909">Histórico de atualizações do Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="38f8d-909">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="38f8d-910">Histórico de atualizações do Office 2016 e 2019 (Clique para Executar)</span><span class="sxs-lookup"><span data-stu-id="38f8d-910">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="38f8d-911">Histórico de atualizações do Office 2013 (clique para executar)</span><span class="sxs-lookup"><span data-stu-id="38f8d-911">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="38f8d-912">Histórico de atualizações do Office 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="38f8d-912">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="38f8d-913">Histórico de atualizações do Outlook 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="38f8d-913">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="38f8d-914">Histórico de atualizações do Office para Mac</span><span class="sxs-lookup"><span data-stu-id="38f8d-914">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
