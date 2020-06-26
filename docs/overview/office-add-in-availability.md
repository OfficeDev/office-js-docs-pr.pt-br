---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos com suporte para o Excel, OneNote, Outlook, PowerPoint, Project e Word.
ms.date: 06/23/2020
localization_priority: Priority
ms.openlocfilehash: 979c873b1c5f2d1d7847414f037d5c75737aa33d
ms.sourcegitcommit: a4873c3525c7d30ef551545d27eb2c0a16b4eb50
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/25/2020
ms.locfileid: "44888156"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="da251-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="da251-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="da251-104">Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="da251-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="da251-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que são compatíveis atualmente com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="da251-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="da251-106">A versão inicial do Office 2016 instalada por meio do MSI apenas contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="da251-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="da251-107">Para saber mais sobre o histórico de atualizações de várias versões do Office, confira a seção[Confira também](#see-also).</span><span class="sxs-lookup"><span data-stu-id="da251-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="da251-108">Excel</span><span class="sxs-lookup"><span data-stu-id="da251-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="da251-109">Plataforma</span><span class="sxs-lookup"><span data-stu-id="da251-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="da251-110">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="da251-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="da251-111">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="da251-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="da251-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="da251-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-113">Office na Web</span><span class="sxs-lookup"><span data-stu-id="da251-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="da251-114">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="da251-114">- TaskPane</span></span><br><span data-ttu-id="da251-115">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="da251-115">
        - Content</span></span><br><span data-ttu-id="da251-116">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="da251-116">
        - Custom Functions</span></span><br><span data-ttu-id="da251-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="da251-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="da251-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="da251-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="da251-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="da251-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="da251-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="da251-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="da251-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="da251-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="da251-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="da251-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="da251-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="da251-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="da251-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="da251-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="da251-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="da251-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="da251-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="da251-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="da251-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="da251-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="da251-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="da251-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="da251-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="da251-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="da251-131">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="da251-131">
        - BindingEvents</span></span><br><span data-ttu-id="da251-132">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da251-132">
        - CompressedFile</span></span><br><span data-ttu-id="da251-133">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da251-133">
        - DocumentEvents</span></span><br><span data-ttu-id="da251-134">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="da251-134">
        - File</span></span><br><span data-ttu-id="da251-135">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="da251-135">
        - MatrixBindings</span></span><br><span data-ttu-id="da251-136">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-136">
        - MatrixCoercion</span></span><br><span data-ttu-id="da251-137">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="da251-137">
        - Selection</span></span><br><span data-ttu-id="da251-138">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="da251-138">
        - Settings</span></span><br><span data-ttu-id="da251-139">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="da251-139">
        - TableBindings</span></span><br><span data-ttu-id="da251-140">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-140">
        - TableCoercion</span></span><br><span data-ttu-id="da251-141">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="da251-141">
        - TextBindings</span></span><br><span data-ttu-id="da251-142">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-142">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-143">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="da251-143">Office on Windows</span></span><br><span data-ttu-id="da251-144">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="da251-144">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="da251-145">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="da251-145">- TaskPane</span></span><br><span data-ttu-id="da251-146">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="da251-146">
        - Content</span></span><br><span data-ttu-id="da251-147">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="da251-147">
        - Custom Functions</span></span><br><span data-ttu-id="da251-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="da251-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="da251-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="da251-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="da251-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="da251-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="da251-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="da251-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="da251-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="da251-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="da251-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="da251-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="da251-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="da251-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="da251-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="da251-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="da251-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="da251-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="da251-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="da251-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="da251-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="da251-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="da251-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="da251-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="da251-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="da251-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="da251-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="da251-163">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="da251-163">
        - BindingEvents</span></span><br><span data-ttu-id="da251-164">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da251-164">
        - CompressedFile</span></span><br><span data-ttu-id="da251-165">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da251-165">
        - DocumentEvents</span></span><br><span data-ttu-id="da251-166">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="da251-166">
        - File</span></span><br><span data-ttu-id="da251-167">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="da251-167">
        - MatrixBindings</span></span><br><span data-ttu-id="da251-168">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-168">
        - MatrixCoercion</span></span><br><span data-ttu-id="da251-169">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="da251-169">
        - Selection</span></span><br><span data-ttu-id="da251-170">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="da251-170">
        - Settings</span></span><br><span data-ttu-id="da251-171">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="da251-171">
        - TableBindings</span></span><br><span data-ttu-id="da251-172">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-172">
        - TableCoercion</span></span><br><span data-ttu-id="da251-173">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="da251-173">
        - TextBindings</span></span><br><span data-ttu-id="da251-174">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-174">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-175">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="da251-175">Office 2019 on Windows</span></span><br><span data-ttu-id="da251-176">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="da251-176">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="da251-177">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="da251-177">- TaskPane</span></span><br><span data-ttu-id="da251-178">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="da251-178">
        - Content</span></span><br><span data-ttu-id="da251-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="da251-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="da251-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="da251-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="da251-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="da251-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="da251-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="da251-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="da251-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="da251-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="da251-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="da251-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="da251-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="da251-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="da251-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="da251-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="da251-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="da251-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="da251-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="da251-190">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="da251-190">- BindingEvents</span></span><br><span data-ttu-id="da251-191">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da251-191">
        - CompressedFile</span></span><br><span data-ttu-id="da251-192">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da251-192">
        - DocumentEvents</span></span><br><span data-ttu-id="da251-193">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="da251-193">
        - File</span></span><br><span data-ttu-id="da251-194">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="da251-194">
        - MatrixBindings</span></span><br><span data-ttu-id="da251-195">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-195">
        - MatrixCoercion</span></span><br><span data-ttu-id="da251-196">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="da251-196">
        - Selection</span></span><br><span data-ttu-id="da251-197">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="da251-197">
        - Settings</span></span><br><span data-ttu-id="da251-198">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="da251-198">
        - TableBindings</span></span><br><span data-ttu-id="da251-199">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-199">
        - TableCoercion</span></span><br><span data-ttu-id="da251-200">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="da251-200">
        - TextBindings</span></span><br><span data-ttu-id="da251-201">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-201">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-202">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="da251-202">Office 2016 on Windows</span></span><br><span data-ttu-id="da251-203">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="da251-203">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="da251-204">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="da251-204">- TaskPane</span></span><br><span data-ttu-id="da251-205">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="da251-205">
        - Content</span></span></td>
    <td><span data-ttu-id="da251-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="da251-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="da251-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="da251-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="da251-209">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="da251-209">- BindingEvents</span></span><br><span data-ttu-id="da251-210">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da251-210">
        - CompressedFile</span></span><br><span data-ttu-id="da251-211">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da251-211">
        - DocumentEvents</span></span><br><span data-ttu-id="da251-212">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="da251-212">
        - File</span></span><br><span data-ttu-id="da251-213">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="da251-213">
        - MatrixBindings</span></span><br><span data-ttu-id="da251-214">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-214">
        - MatrixCoercion</span></span><br><span data-ttu-id="da251-215">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="da251-215">
        - Selection</span></span><br><span data-ttu-id="da251-216">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="da251-216">
        - Settings</span></span><br><span data-ttu-id="da251-217">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="da251-217">
        - TableBindings</span></span><br><span data-ttu-id="da251-218">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-218">
        - TableCoercion</span></span><br><span data-ttu-id="da251-219">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="da251-219">
        - TextBindings</span></span><br><span data-ttu-id="da251-220">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-220">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-221">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="da251-221">Office 2013 on Windows</span></span><br><span data-ttu-id="da251-222">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="da251-222">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="da251-223">
        - TaskPane</span><span class="sxs-lookup"><span data-stu-id="da251-223">
        - TaskPane</span></span><br><span data-ttu-id="da251-224">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="da251-224">
        - Content</span></span></td>
    <td>  <span data-ttu-id="da251-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="da251-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="da251-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="da251-227">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="da251-227">
        - BindingEvents</span></span><br><span data-ttu-id="da251-228">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da251-228">
        - DocumentEvents</span></span><br><span data-ttu-id="da251-229">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="da251-229">
        - File</span></span><br><span data-ttu-id="da251-230">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="da251-230">
        - MatrixBindings</span></span><br><span data-ttu-id="da251-231">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-231">
        - MatrixCoercion</span></span><br><span data-ttu-id="da251-232">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="da251-232">
        - Selection</span></span><br><span data-ttu-id="da251-233">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="da251-233">
        - Settings</span></span><br><span data-ttu-id="da251-234">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="da251-234">
        - TableBindings</span></span><br><span data-ttu-id="da251-235">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-235">
        - TableCoercion</span></span><br><span data-ttu-id="da251-236">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="da251-236">
        - TextBindings</span></span><br><span data-ttu-id="da251-237">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-237">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-238">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="da251-238">Office on iPad</span></span><br><span data-ttu-id="da251-239">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="da251-239">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="da251-240">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="da251-240">- TaskPane</span></span><br><span data-ttu-id="da251-241">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="da251-241">
        - Content</span></span></td>
    <td><span data-ttu-id="da251-242">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-242">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="da251-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="da251-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="da251-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="da251-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="da251-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="da251-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="da251-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="da251-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="da251-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="da251-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="da251-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="da251-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="da251-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="da251-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="da251-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="da251-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="da251-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="da251-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="da251-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="da251-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="da251-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="da251-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="da251-255">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="da251-255">- BindingEvents</span></span><br><span data-ttu-id="da251-256">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da251-256">
        - DocumentEvents</span></span><br><span data-ttu-id="da251-257">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="da251-257">
        - File</span></span><br><span data-ttu-id="da251-258">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="da251-258">
        - MatrixBindings</span></span><br><span data-ttu-id="da251-259">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-259">
        - MatrixCoercion</span></span><br><span data-ttu-id="da251-260">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="da251-260">
        - Selection</span></span><br><span data-ttu-id="da251-261">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="da251-261">
        - Settings</span></span><br><span data-ttu-id="da251-262">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="da251-262">
        - TableBindings</span></span><br><span data-ttu-id="da251-263">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-263">
        - TableCoercion</span></span><br><span data-ttu-id="da251-264">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="da251-264">
        - TextBindings</span></span><br><span data-ttu-id="da251-265">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-265">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-266">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="da251-266">Office on Mac</span></span><br><span data-ttu-id="da251-267">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="da251-267">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="da251-268">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="da251-268">- TaskPane</span></span><br><span data-ttu-id="da251-269">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="da251-269">
        - Content</span></span><br><span data-ttu-id="da251-270">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="da251-270">
        - Custom Functions</span></span><br><span data-ttu-id="da251-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="da251-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="da251-272">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-272">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="da251-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="da251-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="da251-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="da251-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="da251-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="da251-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="da251-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="da251-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="da251-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="da251-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="da251-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="da251-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="da251-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="da251-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="da251-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="da251-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="da251-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="da251-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="da251-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="da251-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="da251-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="da251-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="da251-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="da251-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="da251-286">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="da251-286">- BindingEvents</span></span><br><span data-ttu-id="da251-287">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da251-287">
        - CompressedFile</span></span><br><span data-ttu-id="da251-288">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da251-288">
        - DocumentEvents</span></span><br><span data-ttu-id="da251-289">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="da251-289">
        - File</span></span><br><span data-ttu-id="da251-290">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="da251-290">
        - MatrixBindings</span></span><br><span data-ttu-id="da251-291">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-291">
        - MatrixCoercion</span></span><br><span data-ttu-id="da251-292">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="da251-292">
        - PdfFile</span></span><br><span data-ttu-id="da251-293">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="da251-293">
        - Selection</span></span><br><span data-ttu-id="da251-294">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="da251-294">
        - Settings</span></span><br><span data-ttu-id="da251-295">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="da251-295">
        - TableBindings</span></span><br><span data-ttu-id="da251-296">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-296">
        - TableCoercion</span></span><br><span data-ttu-id="da251-297">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="da251-297">
        - TextBindings</span></span><br><span data-ttu-id="da251-298">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-298">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-299">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="da251-299">Office 2019 on Mac</span></span><br><span data-ttu-id="da251-300">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="da251-300">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="da251-301">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="da251-301">- TaskPane</span></span><br><span data-ttu-id="da251-302">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="da251-302">
        - Content</span></span><br><span data-ttu-id="da251-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="da251-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="da251-304">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-304">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="da251-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="da251-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="da251-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="da251-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="da251-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="da251-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="da251-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="da251-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="da251-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="da251-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="da251-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="da251-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="da251-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="da251-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="da251-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="da251-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="da251-314">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="da251-314">- BindingEvents</span></span><br><span data-ttu-id="da251-315">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da251-315">
        - CompressedFile</span></span><br><span data-ttu-id="da251-316">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da251-316">
        - DocumentEvents</span></span><br><span data-ttu-id="da251-317">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="da251-317">
        - File</span></span><br><span data-ttu-id="da251-318">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="da251-318">
        - MatrixBindings</span></span><br><span data-ttu-id="da251-319">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-319">
        - MatrixCoercion</span></span><br><span data-ttu-id="da251-320">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="da251-320">
        - PdfFile</span></span><br><span data-ttu-id="da251-321">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="da251-321">
        - Selection</span></span><br><span data-ttu-id="da251-322">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="da251-322">
        - Settings</span></span><br><span data-ttu-id="da251-323">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="da251-323">
        - TableBindings</span></span><br><span data-ttu-id="da251-324">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-324">
        - TableCoercion</span></span><br><span data-ttu-id="da251-325">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="da251-325">
        - TextBindings</span></span><br><span data-ttu-id="da251-326">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-326">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-327">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="da251-327">Office 2016 on Mac</span></span><br><span data-ttu-id="da251-328">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="da251-328">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="da251-329">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="da251-329">- TaskPane</span></span><br><span data-ttu-id="da251-330">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="da251-330">
        - Content</span></span></td>
    <td><span data-ttu-id="da251-331">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-331">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="da251-332">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="da251-332">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="da251-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="da251-334">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="da251-334">- BindingEvents</span></span><br><span data-ttu-id="da251-335">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da251-335">
        - CompressedFile</span></span><br><span data-ttu-id="da251-336">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da251-336">
        - DocumentEvents</span></span><br><span data-ttu-id="da251-337">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="da251-337">
        - File</span></span><br><span data-ttu-id="da251-338">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="da251-338">
        - MatrixBindings</span></span><br><span data-ttu-id="da251-339">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-339">
        - MatrixCoercion</span></span><br><span data-ttu-id="da251-340">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="da251-340">
        - PdfFile</span></span><br><span data-ttu-id="da251-341">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="da251-341">
        - Selection</span></span><br><span data-ttu-id="da251-342">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="da251-342">
        - Settings</span></span><br><span data-ttu-id="da251-343">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="da251-343">
        - TableBindings</span></span><br><span data-ttu-id="da251-344">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-344">
        - TableCoercion</span></span><br><span data-ttu-id="da251-345">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="da251-345">
        - TextBindings</span></span><br><span data-ttu-id="da251-346">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-346">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="da251-347">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="da251-347">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="da251-348">Funções personalizadas (somente Excel)</span><span class="sxs-lookup"><span data-stu-id="da251-348">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="da251-349">Plataforma</span><span class="sxs-lookup"><span data-stu-id="da251-349">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="da251-350">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="da251-350">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="da251-351">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="da251-351">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="da251-352"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="da251-352"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-353">Office na Web</span><span class="sxs-lookup"><span data-stu-id="da251-353">Office on the web</span></span></td>
    <td><span data-ttu-id="da251-354">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="da251-354">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="da251-355">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-355">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-356">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="da251-356">Office on Windows</span></span><br><span data-ttu-id="da251-357">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="da251-357">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="da251-358">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="da251-358">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="da251-359">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-359">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-360">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="da251-360">Office on Mac</span></span><br><span data-ttu-id="da251-361">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="da251-361">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="da251-362">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="da251-362">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="da251-363">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-363">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="da251-364">Outlook</span><span class="sxs-lookup"><span data-stu-id="da251-364">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="da251-365">Plataforma</span><span class="sxs-lookup"><span data-stu-id="da251-365">Platform</span></span></th>
    <th><span data-ttu-id="da251-366">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="da251-366">Extension points</span></span></th>
    <th><span data-ttu-id="da251-367">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="da251-367">API requirement sets</span></span></th>
    <th><span data-ttu-id="da251-368"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="da251-368"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-369">Office na Web</span><span class="sxs-lookup"><span data-stu-id="da251-369">Office on the web</span></span><br><span data-ttu-id="da251-370">(moderno)</span><span class="sxs-lookup"><span data-stu-id="da251-370">(modern)</span></span></td>
    <td> <span data-ttu-id="da251-371">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="da251-371">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="da251-372">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="da251-372">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="da251-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="da251-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="da251-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="da251-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="da251-375">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="da251-375">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="da251-376">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-376">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="da251-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="da251-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="da251-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="da251-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="da251-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="da251-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="da251-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="da251-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="da251-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="da251-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="da251-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="da251-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="da251-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a></span><span class="sxs-lookup"><span data-stu-id="da251-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="da251-384">Não disponível</span><span class="sxs-lookup"><span data-stu-id="da251-384">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-385">Office na Web</span><span class="sxs-lookup"><span data-stu-id="da251-385">Office on the web</span></span><br><span data-ttu-id="da251-386">(clássico)</span><span class="sxs-lookup"><span data-stu-id="da251-386">(classic)</span></span></td>
    <td> <span data-ttu-id="da251-387">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="da251-387">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="da251-388">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="da251-388">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="da251-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="da251-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="da251-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="da251-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="da251-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="da251-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="da251-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="da251-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="da251-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="da251-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="da251-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="da251-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="da251-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="da251-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="da251-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="da251-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="da251-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="da251-398">Não disponível</span><span class="sxs-lookup"><span data-stu-id="da251-398">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-399">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="da251-399">Office on Windows</span></span><br><span data-ttu-id="da251-400">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="da251-400">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="da251-401">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="da251-401">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="da251-402">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="da251-402">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="da251-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="da251-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="da251-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="da251-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="da251-405">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="da251-405">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="da251-406">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Módulos</a></span><span class="sxs-lookup"><span data-stu-id="da251-406">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="da251-407">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-407">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="da251-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="da251-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="da251-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="da251-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="da251-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="da251-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="da251-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="da251-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="da251-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="da251-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="da251-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="da251-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="da251-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a></span><span class="sxs-lookup"><span data-stu-id="da251-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="da251-415">Não disponível</span><span class="sxs-lookup"><span data-stu-id="da251-415">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-416">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="da251-416">Office 2019 on Windows</span></span><br><span data-ttu-id="da251-417">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="da251-417">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="da251-418">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="da251-418">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="da251-419">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="da251-419">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="da251-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="da251-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="da251-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="da251-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="da251-422">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="da251-422">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="da251-423">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Módulos</a></span><span class="sxs-lookup"><span data-stu-id="da251-423">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="da251-424">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-424">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="da251-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="da251-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="da251-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="da251-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="da251-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="da251-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="da251-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="da251-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="da251-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="da251-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="da251-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="da251-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="da251-431">Não disponível</span><span class="sxs-lookup"><span data-stu-id="da251-431">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-432">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="da251-432">Office 2016 on Windows</span></span><br><span data-ttu-id="da251-433">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="da251-433">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="da251-434">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="da251-434">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="da251-435">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="da251-435">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="da251-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="da251-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="da251-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="da251-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="da251-438">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="da251-438">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="da251-439">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Módulos</a></span><span class="sxs-lookup"><span data-stu-id="da251-439">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="da251-440">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-440">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="da251-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="da251-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="da251-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="da251-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="da251-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="da251-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="da251-444">Não disponível</span><span class="sxs-lookup"><span data-stu-id="da251-444">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-445">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="da251-445">Office 2013 on Windows</span></span><br><span data-ttu-id="da251-446">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="da251-446">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="da251-447">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="da251-447">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="da251-448">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="da251-448">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="da251-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="da251-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="da251-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="da251-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br>
    <td> <span data-ttu-id="da251-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="da251-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="da251-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="da251-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="da251-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="da251-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="da251-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="da251-455">Não disponível</span><span class="sxs-lookup"><span data-stu-id="da251-455">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-456">Office no iOS</span><span class="sxs-lookup"><span data-stu-id="da251-456">Office on iOS</span></span><br><span data-ttu-id="da251-457">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="da251-457">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="da251-458">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="da251-458">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="da251-459">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="da251-459">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="da251-460">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-460">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="da251-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="da251-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="da251-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="da251-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="da251-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="da251-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="da251-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="da251-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="da251-465">Não disponível</span><span class="sxs-lookup"><span data-stu-id="da251-465">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-466">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="da251-466">Office on Mac</span></span><br><span data-ttu-id="da251-467">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="da251-467">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="da251-468">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="da251-468">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="da251-469">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="da251-469">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="da251-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="da251-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="da251-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="da251-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="da251-472">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="da251-472">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="da251-473">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-473">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="da251-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="da251-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="da251-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="da251-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="da251-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="da251-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="da251-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="da251-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="da251-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="da251-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="da251-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="da251-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="da251-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a></span><span class="sxs-lookup"><span data-stu-id="da251-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="da251-481">Não disponível</span><span class="sxs-lookup"><span data-stu-id="da251-481">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-482">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="da251-482">Office 2019 on Mac</span></span><br><span data-ttu-id="da251-483">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="da251-483">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="da251-484">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="da251-484">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="da251-485">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="da251-485">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="da251-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="da251-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="da251-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="da251-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="da251-488">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="da251-488">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="da251-489">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-489">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="da251-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="da251-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="da251-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="da251-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="da251-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="da251-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="da251-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="da251-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="da251-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="da251-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="da251-495">Não disponível</span><span class="sxs-lookup"><span data-stu-id="da251-495">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-496">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="da251-496">Office 2016 on Mac</span></span><br><span data-ttu-id="da251-497">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="da251-497">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="da251-498">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="da251-498">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="da251-499">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="da251-499">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="da251-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="da251-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="da251-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="da251-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="da251-502">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="da251-502">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="da251-503">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-503">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="da251-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="da251-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="da251-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="da251-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="da251-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="da251-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="da251-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="da251-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="da251-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="da251-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="da251-509">Não disponível</span><span class="sxs-lookup"><span data-stu-id="da251-509">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-510">Outlook no Android</span><span class="sxs-lookup"><span data-stu-id="da251-510">Office on Android</span></span><br><span data-ttu-id="da251-511">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="da251-511">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="da251-512">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="da251-512">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="da251-513">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Organizador de compromissos (Redigir): reunião on-line (visualização)</span><span class="sxs-lookup"><span data-stu-id="da251-513">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Appointment Organizer (Compose): online meeting</a> (preview)</span></span><br><span data-ttu-id="da251-514">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="da251-514">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="da251-515">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-515">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="da251-516">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="da251-516">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="da251-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="da251-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="da251-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="da251-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="da251-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="da251-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="da251-520">Não disponível</span><span class="sxs-lookup"><span data-stu-id="da251-520">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="da251-521">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="da251-521">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="da251-522">O suporte ao cliente para um conjunto de requisitos pode ser restringido pelo suporte do servidor Exchange.</span><span class="sxs-lookup"><span data-stu-id="da251-522">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="da251-523">Consulte [Conjuntos de requisitos da API JavaScript do Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) para obter detalhes sobre o intervalo de conjuntos de requisitos suportado pelo servidor Exchange e pelos clientes Outlook.</span><span class="sxs-lookup"><span data-stu-id="da251-523">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="da251-524">Word</span><span class="sxs-lookup"><span data-stu-id="da251-524">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="da251-525">Plataforma</span><span class="sxs-lookup"><span data-stu-id="da251-525">Platform</span></span></th>
    <th><span data-ttu-id="da251-526">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="da251-526">Extension points</span></span></th>
    <th><span data-ttu-id="da251-527">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="da251-527">API requirement sets</span></span></th>
    <th><span data-ttu-id="da251-528"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="da251-528"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-529">Office na Web</span><span class="sxs-lookup"><span data-stu-id="da251-529">Office on the web</span></span></td>
    <td> <span data-ttu-id="da251-530">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="da251-530">- TaskPane</span></span><br><span data-ttu-id="da251-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="da251-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="da251-532">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-532">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="da251-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="da251-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="da251-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="da251-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="da251-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="da251-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="da251-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="da251-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="da251-538">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="da251-538">- BindingEvents</span></span><br><span data-ttu-id="da251-539">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="da251-539">
         - CustomXmlParts</span></span><br><span data-ttu-id="da251-540">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da251-540">
         - DocumentEvents</span></span><br><span data-ttu-id="da251-541">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="da251-541">
         - File</span></span><br><span data-ttu-id="da251-542">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-542">
         - HtmlCoercion</span></span><br><span data-ttu-id="da251-543">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="da251-543">
         - MatrixBindings</span></span><br><span data-ttu-id="da251-544">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-544">
         - MatrixCoercion</span></span><br><span data-ttu-id="da251-545">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-545">
         - OoxmlCoercion</span></span><br><span data-ttu-id="da251-546">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="da251-546">
         - PdfFile</span></span><br><span data-ttu-id="da251-547">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="da251-547">
         - Selection</span></span><br><span data-ttu-id="da251-548">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="da251-548">
         - Settings</span></span><br><span data-ttu-id="da251-549">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="da251-549">
         - TableBindings</span></span><br><span data-ttu-id="da251-550">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-550">
         - TableCoercion</span></span><br><span data-ttu-id="da251-551">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="da251-551">
         - TextBindings</span></span><br><span data-ttu-id="da251-552">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-552">
         - TextCoercion</span></span><br><span data-ttu-id="da251-553">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="da251-553">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-554">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="da251-554">Office on Windows</span></span><br><span data-ttu-id="da251-555">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="da251-555">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="da251-556">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="da251-556">- TaskPane</span></span><br><span data-ttu-id="da251-557">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="da251-557">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="da251-558">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-558">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="da251-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="da251-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="da251-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="da251-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="da251-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="da251-562">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-562">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="da251-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="da251-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="da251-564">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="da251-564">- BindingEvents</span></span><br><span data-ttu-id="da251-565">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da251-565">
         - CompressedFile</span></span><br><span data-ttu-id="da251-566">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="da251-566">
         - CustomXmlParts</span></span><br><span data-ttu-id="da251-567">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da251-567">
         - DocumentEvents</span></span><br><span data-ttu-id="da251-568">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="da251-568">
         - File</span></span><br><span data-ttu-id="da251-569">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-569">
         - HtmlCoercion</span></span><br><span data-ttu-id="da251-570">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="da251-570">
         - MatrixBindings</span></span><br><span data-ttu-id="da251-571">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-571">
         - MatrixCoercion</span></span><br><span data-ttu-id="da251-572">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-572">
         - OoxmlCoercion</span></span><br><span data-ttu-id="da251-573">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="da251-573">
         - PdfFile</span></span><br><span data-ttu-id="da251-574">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="da251-574">
         - Selection</span></span><br><span data-ttu-id="da251-575">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="da251-575">
         - Settings</span></span><br><span data-ttu-id="da251-576">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="da251-576">
         - TableBindings</span></span><br><span data-ttu-id="da251-577">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-577">
         - TableCoercion</span></span><br><span data-ttu-id="da251-578">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="da251-578">
         - TextBindings</span></span><br><span data-ttu-id="da251-579">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-579">
         - TextCoercion</span></span><br><span data-ttu-id="da251-580">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="da251-580">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-581">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="da251-581">Office 2019 on Windows</span></span><br><span data-ttu-id="da251-582">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="da251-582">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="da251-583">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="da251-583">- TaskPane</span></span><br><span data-ttu-id="da251-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="da251-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="da251-585">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-585">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="da251-586">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="da251-586">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="da251-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="da251-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="da251-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="da251-589">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-589">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="da251-590">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="da251-590">- BindingEvents</span></span><br><span data-ttu-id="da251-591">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da251-591">
         - CompressedFile</span></span><br><span data-ttu-id="da251-592">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="da251-592">
         - CustomXmlParts</span></span><br><span data-ttu-id="da251-593">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da251-593">
         - DocumentEvents</span></span><br><span data-ttu-id="da251-594">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="da251-594">
         - File</span></span><br><span data-ttu-id="da251-595">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-595">
         - HtmlCoercion</span></span><br><span data-ttu-id="da251-596">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="da251-596">
         - MatrixBindings</span></span><br><span data-ttu-id="da251-597">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-597">
         - MatrixCoercion</span></span><br><span data-ttu-id="da251-598">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-598">
         - OoxmlCoercion</span></span><br><span data-ttu-id="da251-599">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="da251-599">
         - PdfFile</span></span><br><span data-ttu-id="da251-600">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="da251-600">
         - Selection</span></span><br><span data-ttu-id="da251-601">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="da251-601">
         - Settings</span></span><br><span data-ttu-id="da251-602">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="da251-602">
         - TableBindings</span></span><br><span data-ttu-id="da251-603">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-603">
         - TableCoercion</span></span><br><span data-ttu-id="da251-604">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="da251-604">
         - TextBindings</span></span><br><span data-ttu-id="da251-605">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-605">
         - TextCoercion</span></span><br><span data-ttu-id="da251-606">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="da251-606">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-607">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="da251-607">Office 2016 on Windows</span></span><br><span data-ttu-id="da251-608">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="da251-608">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="da251-609">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="da251-609">- TaskPane</span></span></td>
    <td> <span data-ttu-id="da251-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="da251-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="da251-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="da251-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="da251-613">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="da251-613">- BindingEvents</span></span><br><span data-ttu-id="da251-614">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da251-614">
         - CompressedFile</span></span><br><span data-ttu-id="da251-615">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="da251-615">
         - CustomXmlParts</span></span><br><span data-ttu-id="da251-616">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da251-616">
         - DocumentEvents</span></span><br><span data-ttu-id="da251-617">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="da251-617">
         - File</span></span><br><span data-ttu-id="da251-618">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-618">
         - HtmlCoercion</span></span><br><span data-ttu-id="da251-619">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="da251-619">
         - MatrixBindings</span></span><br><span data-ttu-id="da251-620">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-620">
         - MatrixCoercion</span></span><br><span data-ttu-id="da251-621">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-621">
         - OoxmlCoercion</span></span><br><span data-ttu-id="da251-622">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="da251-622">
         - PdfFile</span></span><br><span data-ttu-id="da251-623">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="da251-623">
         - Selection</span></span><br><span data-ttu-id="da251-624">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="da251-624">
         - Settings</span></span><br><span data-ttu-id="da251-625">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="da251-625">
         - TableBindings</span></span><br><span data-ttu-id="da251-626">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-626">
         - TableCoercion</span></span><br><span data-ttu-id="da251-627">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="da251-627">
         - TextBindings</span></span><br><span data-ttu-id="da251-628">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-628">
         - TextCoercion</span></span><br><span data-ttu-id="da251-629">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="da251-629">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-630">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="da251-630">Office 2013 on Windows</span></span><br><span data-ttu-id="da251-631">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="da251-631">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="da251-632">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="da251-632">- TaskPane</span></span></td>
    <td> <span data-ttu-id="da251-633">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="da251-633">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="da251-634">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-634">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="da251-635">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="da251-635">- BindingEvents</span></span><br><span data-ttu-id="da251-636">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da251-636">
         - CompressedFile</span></span><br><span data-ttu-id="da251-637">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="da251-637">
         - CustomXmlParts</span></span><br><span data-ttu-id="da251-638">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da251-638">
         - DocumentEvents</span></span><br><span data-ttu-id="da251-639">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="da251-639">
         - File</span></span><br><span data-ttu-id="da251-640">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-640">
         - HtmlCoercion</span></span><br><span data-ttu-id="da251-641">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="da251-641">
         - MatrixBindings</span></span><br><span data-ttu-id="da251-642">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-642">
         - MatrixCoercion</span></span><br><span data-ttu-id="da251-643">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-643">
         - OoxmlCoercion</span></span><br><span data-ttu-id="da251-644">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="da251-644">
         - PdfFile</span></span><br><span data-ttu-id="da251-645">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="da251-645">
         - Selection</span></span><br><span data-ttu-id="da251-646">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="da251-646">
         - Settings</span></span><br><span data-ttu-id="da251-647">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="da251-647">
         - TableBindings</span></span><br><span data-ttu-id="da251-648">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-648">
         - TableCoercion</span></span><br><span data-ttu-id="da251-649">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="da251-649">
         - TextBindings</span></span><br><span data-ttu-id="da251-650">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-650">
         - TextCoercion</span></span><br><span data-ttu-id="da251-651">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="da251-651">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-652">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="da251-652">Office on iPad</span></span><br><span data-ttu-id="da251-653">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="da251-653">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="da251-654">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="da251-654">- TaskPane</span></span></td>
    <td> <span data-ttu-id="da251-655">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-655">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="da251-656">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="da251-656">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="da251-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="da251-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="da251-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="da251-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="da251-660">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="da251-660">- BindingEvents</span></span><br><span data-ttu-id="da251-661">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da251-661">
         - CompressedFile</span></span><br><span data-ttu-id="da251-662">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="da251-662">
         - CustomXmlParts</span></span><br><span data-ttu-id="da251-663">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da251-663">
         - DocumentEvents</span></span><br><span data-ttu-id="da251-664">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="da251-664">
         - File</span></span><br><span data-ttu-id="da251-665">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-665">
         - HtmlCoercion</span></span><br><span data-ttu-id="da251-666">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="da251-666">
         - MatrixBindings</span></span><br><span data-ttu-id="da251-667">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-667">
         - MatrixCoercion</span></span><br><span data-ttu-id="da251-668">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-668">
         - OoxmlCoercion</span></span><br><span data-ttu-id="da251-669">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="da251-669">
         - PdfFile</span></span><br><span data-ttu-id="da251-670">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="da251-670">
         - Selection</span></span><br><span data-ttu-id="da251-671">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="da251-671">
         - Settings</span></span><br><span data-ttu-id="da251-672">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="da251-672">
         - TableBindings</span></span><br><span data-ttu-id="da251-673">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-673">
         - TableCoercion</span></span><br><span data-ttu-id="da251-674">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="da251-674">
         - TextBindings</span></span><br><span data-ttu-id="da251-675">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-675">
         - TextCoercion</span></span><br><span data-ttu-id="da251-676">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="da251-676">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-677">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="da251-677">Office on Mac</span></span><br><span data-ttu-id="da251-678">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="da251-678">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="da251-679">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="da251-679">- TaskPane</span></span><br><span data-ttu-id="da251-680">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="da251-680">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="da251-681">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-681">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="da251-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="da251-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="da251-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="da251-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="da251-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="da251-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="da251-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="da251-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="da251-687">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="da251-687">- BindingEvents</span></span><br><span data-ttu-id="da251-688">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da251-688">
         - CompressedFile</span></span><br><span data-ttu-id="da251-689">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="da251-689">
         - CustomXmlParts</span></span><br><span data-ttu-id="da251-690">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da251-690">
         - DocumentEvents</span></span><br><span data-ttu-id="da251-691">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="da251-691">
         - File</span></span><br><span data-ttu-id="da251-692">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-692">
         - HtmlCoercion</span></span><br><span data-ttu-id="da251-693">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="da251-693">
         - MatrixBindings</span></span><br><span data-ttu-id="da251-694">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-694">
         - MatrixCoercion</span></span><br><span data-ttu-id="da251-695">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-695">
         - OoxmlCoercion</span></span><br><span data-ttu-id="da251-696">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="da251-696">
         - PdfFile</span></span><br><span data-ttu-id="da251-697">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="da251-697">
         - Selection</span></span><br><span data-ttu-id="da251-698">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="da251-698">
         - Settings</span></span><br><span data-ttu-id="da251-699">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="da251-699">
         - TableBindings</span></span><br><span data-ttu-id="da251-700">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-700">
         - TableCoercion</span></span><br><span data-ttu-id="da251-701">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="da251-701">
         - TextBindings</span></span><br><span data-ttu-id="da251-702">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-702">
         - TextCoercion</span></span><br><span data-ttu-id="da251-703">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="da251-703">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-704">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="da251-704">Office 2019 on Mac</span></span><br><span data-ttu-id="da251-705">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="da251-705">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="da251-706">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="da251-706">- TaskPane</span></span><br><span data-ttu-id="da251-707">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="da251-707">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="da251-708">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-708">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="da251-709">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="da251-709">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="da251-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="da251-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="da251-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="da251-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="da251-713">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="da251-713">- BindingEvents</span></span><br><span data-ttu-id="da251-714">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da251-714">
         - CompressedFile</span></span><br><span data-ttu-id="da251-715">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="da251-715">
         - CustomXmlParts</span></span><br><span data-ttu-id="da251-716">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da251-716">
         - DocumentEvents</span></span><br><span data-ttu-id="da251-717">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="da251-717">
         - File</span></span><br><span data-ttu-id="da251-718">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-718">
         - HtmlCoercion</span></span><br><span data-ttu-id="da251-719">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="da251-719">
         - MatrixBindings</span></span><br><span data-ttu-id="da251-720">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-720">
         - MatrixCoercion</span></span><br><span data-ttu-id="da251-721">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-721">
         - OoxmlCoercion</span></span><br><span data-ttu-id="da251-722">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="da251-722">
         - PdfFile</span></span><br><span data-ttu-id="da251-723">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="da251-723">
         - Selection</span></span><br><span data-ttu-id="da251-724">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="da251-724">
         - Settings</span></span><br><span data-ttu-id="da251-725">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="da251-725">
         - TableBindings</span></span><br><span data-ttu-id="da251-726">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-726">
         - TableCoercion</span></span><br><span data-ttu-id="da251-727">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="da251-727">
         - TextBindings</span></span><br><span data-ttu-id="da251-728">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-728">
         - TextCoercion</span></span><br><span data-ttu-id="da251-729">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="da251-729">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-730">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="da251-730">Office 2016 on Mac</span></span><br><span data-ttu-id="da251-731">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="da251-731">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="da251-732">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="da251-732">- TaskPane</span></span></td>
    <td> <span data-ttu-id="da251-733">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-733">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="da251-734">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="da251-734">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="da251-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="da251-736">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="da251-736">- BindingEvents</span></span><br><span data-ttu-id="da251-737">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da251-737">
         - CompressedFile</span></span><br><span data-ttu-id="da251-738">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="da251-738">
         - CustomXmlParts</span></span><br><span data-ttu-id="da251-739">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da251-739">
         - DocumentEvents</span></span><br><span data-ttu-id="da251-740">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="da251-740">
         - File</span></span><br><span data-ttu-id="da251-741">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-741">
         - HtmlCoercion</span></span><br><span data-ttu-id="da251-742">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="da251-742">
         - MatrixBindings</span></span><br><span data-ttu-id="da251-743">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-743">
         - MatrixCoercion</span></span><br><span data-ttu-id="da251-744">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-744">
         - OoxmlCoercion</span></span><br><span data-ttu-id="da251-745">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="da251-745">
         - PdfFile</span></span><br><span data-ttu-id="da251-746">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="da251-746">
         - Selection</span></span><br><span data-ttu-id="da251-747">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="da251-747">
         - Settings</span></span><br><span data-ttu-id="da251-748">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="da251-748">
         - TableBindings</span></span><br><span data-ttu-id="da251-749">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-749">
         - TableCoercion</span></span><br><span data-ttu-id="da251-750">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="da251-750">
         - TextBindings</span></span><br><span data-ttu-id="da251-751">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-751">
         - TextCoercion</span></span><br><span data-ttu-id="da251-752">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="da251-752">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="da251-753">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="da251-753">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="da251-754">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="da251-754">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="da251-755">Plataforma</span><span class="sxs-lookup"><span data-stu-id="da251-755">Platform</span></span></th>
    <th><span data-ttu-id="da251-756">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="da251-756">Extension points</span></span></th>
    <th><span data-ttu-id="da251-757">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="da251-757">API requirement sets</span></span></th>
    <th><span data-ttu-id="da251-758"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="da251-758"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-759">Office na Web</span><span class="sxs-lookup"><span data-stu-id="da251-759">Office on the web</span></span></td>
    <td> <span data-ttu-id="da251-760">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="da251-760">- Content</span></span><br><span data-ttu-id="da251-761">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="da251-761">
         - TaskPane</span></span><br><span data-ttu-id="da251-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="da251-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="da251-763">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-763">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="da251-764">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-764">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="da251-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="da251-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="da251-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="da251-767">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="da251-767">- ActiveView</span></span><br><span data-ttu-id="da251-768">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da251-768">
         - CompressedFile</span></span><br><span data-ttu-id="da251-769">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da251-769">
         - DocumentEvents</span></span><br><span data-ttu-id="da251-770">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="da251-770">
         - File</span></span><br><span data-ttu-id="da251-771">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="da251-771">
         - PdfFile</span></span><br><span data-ttu-id="da251-772">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="da251-772">
         - Selection</span></span><br><span data-ttu-id="da251-773">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="da251-773">
         - Settings</span></span><br><span data-ttu-id="da251-774">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-774">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-775">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="da251-775">Office on Windows</span></span><br><span data-ttu-id="da251-776">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="da251-776">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="da251-777">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="da251-777">- Content</span></span><br><span data-ttu-id="da251-778">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="da251-778">
         - TaskPane</span></span><br><span data-ttu-id="da251-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="da251-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="da251-780">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-780">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="da251-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="da251-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="da251-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="da251-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="da251-784">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="da251-784">- ActiveView</span></span><br><span data-ttu-id="da251-785">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da251-785">
         - CompressedFile</span></span><br><span data-ttu-id="da251-786">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da251-786">
         - DocumentEvents</span></span><br><span data-ttu-id="da251-787">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="da251-787">
         - File</span></span><br><span data-ttu-id="da251-788">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="da251-788">
         - PdfFile</span></span><br><span data-ttu-id="da251-789">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="da251-789">
         - Selection</span></span><br><span data-ttu-id="da251-790">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="da251-790">
         - Settings</span></span><br><span data-ttu-id="da251-791">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-791">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-792">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="da251-792">Office 2019 on Windows</span></span><br><span data-ttu-id="da251-793">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="da251-793">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="da251-794">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="da251-794">- Content</span></span><br><span data-ttu-id="da251-795">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="da251-795">
         - TaskPane</span></span><br><span data-ttu-id="da251-796">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="da251-796">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="da251-797">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-797">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="da251-798">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-798">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="da251-799">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="da251-799">- ActiveView</span></span><br><span data-ttu-id="da251-800">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da251-800">
         - CompressedFile</span></span><br><span data-ttu-id="da251-801">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da251-801">
         - DocumentEvents</span></span><br><span data-ttu-id="da251-802">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="da251-802">
         - File</span></span><br><span data-ttu-id="da251-803">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="da251-803">
         - PdfFile</span></span><br><span data-ttu-id="da251-804">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="da251-804">
         - Selection</span></span><br><span data-ttu-id="da251-805">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="da251-805">
         - Settings</span></span><br><span data-ttu-id="da251-806">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-806">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-807">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="da251-807">Office 2016 on Windows</span></span><br><span data-ttu-id="da251-808">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="da251-808">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="da251-809">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="da251-809">- Content</span></span><br><span data-ttu-id="da251-810">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="da251-810">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="da251-811">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="da251-811">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="da251-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="da251-813">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="da251-813">- ActiveView</span></span><br><span data-ttu-id="da251-814">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da251-814">
         - CompressedFile</span></span><br><span data-ttu-id="da251-815">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da251-815">
         - DocumentEvents</span></span><br><span data-ttu-id="da251-816">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="da251-816">
         - File</span></span><br><span data-ttu-id="da251-817">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="da251-817">
         - PdfFile</span></span><br><span data-ttu-id="da251-818">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="da251-818">
         - Selection</span></span><br><span data-ttu-id="da251-819">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="da251-819">
         - Settings</span></span><br><span data-ttu-id="da251-820">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-820">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-821">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="da251-821">Office 2013 on Windows</span></span><br><span data-ttu-id="da251-822">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="da251-822">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="da251-823">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="da251-823">- Content</span></span><br><span data-ttu-id="da251-824">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="da251-824">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="da251-825">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="da251-825">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="da251-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="da251-827">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="da251-827">- ActiveView</span></span><br><span data-ttu-id="da251-828">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da251-828">
         - CompressedFile</span></span><br><span data-ttu-id="da251-829">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da251-829">
         - DocumentEvents</span></span><br><span data-ttu-id="da251-830">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="da251-830">
         - File</span></span><br><span data-ttu-id="da251-831">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="da251-831">
         - PdfFile</span></span><br><span data-ttu-id="da251-832">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="da251-832">
         - Selection</span></span><br><span data-ttu-id="da251-833">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="da251-833">
         - Settings</span></span><br><span data-ttu-id="da251-834">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-834">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-835">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="da251-835">Office on iPad</span></span><br><span data-ttu-id="da251-836">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="da251-836">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="da251-837">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="da251-837">- Content</span></span><br><span data-ttu-id="da251-838">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="da251-838">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="da251-839">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-839">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="da251-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="da251-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="da251-842">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="da251-842">- ActiveView</span></span><br><span data-ttu-id="da251-843">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da251-843">
         - CompressedFile</span></span><br><span data-ttu-id="da251-844">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da251-844">
         - DocumentEvents</span></span><br><span data-ttu-id="da251-845">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="da251-845">
         - File</span></span><br><span data-ttu-id="da251-846">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="da251-846">
         - PdfFile</span></span><br><span data-ttu-id="da251-847">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="da251-847">
         - Selection</span></span><br><span data-ttu-id="da251-848">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="da251-848">
         - Settings</span></span><br><span data-ttu-id="da251-849">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-849">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-850">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="da251-850">Office on Mac</span></span><br><span data-ttu-id="da251-851">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="da251-851">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="da251-852">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="da251-852">- Content</span></span><br><span data-ttu-id="da251-853">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="da251-853">
         - TaskPane</span></span><br><span data-ttu-id="da251-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="da251-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="da251-855">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-855">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="da251-856">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-856">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="da251-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="da251-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="da251-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="da251-859">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="da251-859">- ActiveView</span></span><br><span data-ttu-id="da251-860">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da251-860">
         - CompressedFile</span></span><br><span data-ttu-id="da251-861">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da251-861">
         - DocumentEvents</span></span><br><span data-ttu-id="da251-862">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="da251-862">
         - File</span></span><br><span data-ttu-id="da251-863">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="da251-863">
         - PdfFile</span></span><br><span data-ttu-id="da251-864">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="da251-864">
         - Selection</span></span><br><span data-ttu-id="da251-865">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="da251-865">
         - Settings</span></span><br><span data-ttu-id="da251-866">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-866">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-867">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="da251-867">Office 2019 on Mac</span></span><br><span data-ttu-id="da251-868">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="da251-868">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="da251-869">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="da251-869">- Content</span></span><br><span data-ttu-id="da251-870">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="da251-870">
         - TaskPane</span></span><br><span data-ttu-id="da251-871">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="da251-871">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="da251-872">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-872">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="da251-873">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-873">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="da251-874">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="da251-874">- ActiveView</span></span><br><span data-ttu-id="da251-875">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da251-875">
         - CompressedFile</span></span><br><span data-ttu-id="da251-876">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da251-876">
         - DocumentEvents</span></span><br><span data-ttu-id="da251-877">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="da251-877">
         - File</span></span><br><span data-ttu-id="da251-878">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="da251-878">
         - PdfFile</span></span><br><span data-ttu-id="da251-879">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="da251-879">
         - Selection</span></span><br><span data-ttu-id="da251-880">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="da251-880">
         - Settings</span></span><br><span data-ttu-id="da251-881">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-881">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-882">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="da251-882">Office 2016 on Mac</span></span><br><span data-ttu-id="da251-883">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="da251-883">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="da251-884">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="da251-884">- Content</span></span><br><span data-ttu-id="da251-885">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="da251-885">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="da251-886">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="da251-886">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="da251-887">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-887">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="da251-888">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="da251-888">- ActiveView</span></span><br><span data-ttu-id="da251-889">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="da251-889">
         - CompressedFile</span></span><br><span data-ttu-id="da251-890">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da251-890">
         - DocumentEvents</span></span><br><span data-ttu-id="da251-891">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="da251-891">
         - File</span></span><br><span data-ttu-id="da251-892">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="da251-892">
         - PdfFile</span></span><br><span data-ttu-id="da251-893">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="da251-893">
         - Selection</span></span><br><span data-ttu-id="da251-894">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="da251-894">
         - Settings</span></span><br><span data-ttu-id="da251-895">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-895">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="da251-896">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="da251-896">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="da251-897">OneNote</span><span class="sxs-lookup"><span data-stu-id="da251-897">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="da251-898">Plataforma</span><span class="sxs-lookup"><span data-stu-id="da251-898">Platform</span></span></th>
    <th><span data-ttu-id="da251-899">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="da251-899">Extension points</span></span></th>
    <th><span data-ttu-id="da251-900">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="da251-900">API requirement sets</span></span></th>
    <th><span data-ttu-id="da251-901"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="da251-901"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-902">Office na Web</span><span class="sxs-lookup"><span data-stu-id="da251-902">Office on the web</span></span></td>
    <td> <span data-ttu-id="da251-903">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="da251-903">- Content</span></span><br><span data-ttu-id="da251-904">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="da251-904">
         - TaskPane</span></span><br><span data-ttu-id="da251-905">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="da251-905">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="da251-906">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-906">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="da251-907">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-907">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="da251-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="da251-909">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="da251-909">- DocumentEvents</span></span><br><span data-ttu-id="da251-910">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-910">
         - HtmlCoercion</span></span><br><span data-ttu-id="da251-911">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="da251-911">
         - Settings</span></span><br><span data-ttu-id="da251-912">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-912">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="da251-913">Project</span><span class="sxs-lookup"><span data-stu-id="da251-913">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="da251-914">Plataforma</span><span class="sxs-lookup"><span data-stu-id="da251-914">Platform</span></span></th>
    <th><span data-ttu-id="da251-915">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="da251-915">Extension points</span></span></th>
    <th><span data-ttu-id="da251-916">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="da251-916">API requirement sets</span></span></th>
    <th><span data-ttu-id="da251-917"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="da251-917"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-918">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="da251-918">Office 2019 on Windows</span></span><br><span data-ttu-id="da251-919">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="da251-919">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="da251-920">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="da251-920">- TaskPane</span></span></td>
    <td> <span data-ttu-id="da251-921">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-921">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="da251-922">- Seleção</span><span class="sxs-lookup"><span data-stu-id="da251-922">- Selection</span></span><br><span data-ttu-id="da251-923">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-923">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-924">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="da251-924">Office 2016 on Windows</span></span><br><span data-ttu-id="da251-925">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="da251-925">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="da251-926">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="da251-926">- TaskPane</span></span></td>
    <td> <span data-ttu-id="da251-927">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-927">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="da251-928">- Seleção</span><span class="sxs-lookup"><span data-stu-id="da251-928">- Selection</span></span><br><span data-ttu-id="da251-929">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-929">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="da251-930">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="da251-930">Office 2013 on Windows</span></span><br><span data-ttu-id="da251-931">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="da251-931">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="da251-932">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="da251-932">- TaskPane</span></span></td>
    <td> <span data-ttu-id="da251-933">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="da251-933">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="da251-934">- Seleção</span><span class="sxs-lookup"><span data-stu-id="da251-934">- Selection</span></span><br><span data-ttu-id="da251-935">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="da251-935">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="da251-936">Confira também</span><span class="sxs-lookup"><span data-stu-id="da251-936">See also</span></span>

- [<span data-ttu-id="da251-937">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="da251-937">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="da251-938">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="da251-938">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="da251-939">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="da251-939">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="da251-940">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="da251-940">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="da251-941">Documentação de Referência da API</span><span class="sxs-lookup"><span data-stu-id="da251-941">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="da251-942">Histórico de atualizações do Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="da251-942">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="da251-943">Histórico de atualizações do Office 2016 e 2019 (Clique para Executar)</span><span class="sxs-lookup"><span data-stu-id="da251-943">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="da251-944">Histórico de atualizações do Office 2013 (clique para executar)</span><span class="sxs-lookup"><span data-stu-id="da251-944">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="da251-945">Histórico de atualizações do Office 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="da251-945">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="da251-946">Histórico de atualizações do Outlook 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="da251-946">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="da251-947">Histórico de atualizações do Office para Mac</span><span class="sxs-lookup"><span data-stu-id="da251-947">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="da251-948">Criando Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="da251-948">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)