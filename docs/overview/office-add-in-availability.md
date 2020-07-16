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
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="57f70-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="57f70-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="57f70-104">Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="57f70-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="57f70-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que são compatíveis atualmente com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="57f70-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="57f70-106">A versão inicial do Office 2016 instalada por meio do MSI apenas contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="57f70-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="57f70-107">Para saber mais sobre o histórico de atualizações de várias versões do Office, confira a seção[Confira também](#see-also).</span><span class="sxs-lookup"><span data-stu-id="57f70-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="57f70-108">Excel</span><span class="sxs-lookup"><span data-stu-id="57f70-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="57f70-109">Plataforma</span><span class="sxs-lookup"><span data-stu-id="57f70-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="57f70-110">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="57f70-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="57f70-111">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="57f70-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="57f70-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="57f70-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-113">Office na Web</span><span class="sxs-lookup"><span data-stu-id="57f70-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="57f70-114">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="57f70-114">- TaskPane</span></span><br><span data-ttu-id="57f70-115">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="57f70-115">
        - Content</span></span><br><span data-ttu-id="57f70-116">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="57f70-116">
        - Custom Functions</span></span><br><span data-ttu-id="57f70-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="57f70-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="57f70-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="57f70-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="57f70-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="57f70-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="57f70-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="57f70-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="57f70-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="57f70-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="57f70-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="57f70-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="57f70-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="57f70-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="57f70-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="57f70-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="57f70-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="57f70-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="57f70-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="57f70-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="57f70-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="57f70-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="57f70-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="57f70-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="57f70-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="57f70-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="57f70-131">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-131">
        - BindingEvents</span></span><br><span data-ttu-id="57f70-132">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="57f70-132">
        - CompressedFile</span></span><br><span data-ttu-id="57f70-133">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-133">
        - DocumentEvents</span></span><br><span data-ttu-id="57f70-134">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="57f70-134">
        - File</span></span><br><span data-ttu-id="57f70-135">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-135">
        - MatrixBindings</span></span><br><span data-ttu-id="57f70-136">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-136">
        - MatrixCoercion</span></span><br><span data-ttu-id="57f70-137">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="57f70-137">
        - Selection</span></span><br><span data-ttu-id="57f70-138">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="57f70-138">
        - Settings</span></span><br><span data-ttu-id="57f70-139">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-139">
        - TableBindings</span></span><br><span data-ttu-id="57f70-140">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-140">
        - TableCoercion</span></span><br><span data-ttu-id="57f70-141">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-141">
        - TextBindings</span></span><br><span data-ttu-id="57f70-142">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-142">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-143">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="57f70-143">Office on Windows</span></span><br><span data-ttu-id="57f70-144">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="57f70-144">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="57f70-145">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="57f70-145">- TaskPane</span></span><br><span data-ttu-id="57f70-146">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="57f70-146">
        - Content</span></span><br><span data-ttu-id="57f70-147">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="57f70-147">
        - Custom Functions</span></span><br><span data-ttu-id="57f70-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="57f70-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="57f70-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="57f70-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="57f70-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="57f70-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="57f70-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="57f70-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="57f70-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="57f70-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="57f70-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="57f70-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="57f70-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="57f70-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="57f70-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="57f70-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="57f70-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="57f70-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="57f70-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="57f70-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="57f70-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="57f70-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="57f70-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="57f70-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="57f70-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="57f70-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="57f70-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="57f70-163">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-163">
        - BindingEvents</span></span><br><span data-ttu-id="57f70-164">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="57f70-164">
        - CompressedFile</span></span><br><span data-ttu-id="57f70-165">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-165">
        - DocumentEvents</span></span><br><span data-ttu-id="57f70-166">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="57f70-166">
        - File</span></span><br><span data-ttu-id="57f70-167">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-167">
        - MatrixBindings</span></span><br><span data-ttu-id="57f70-168">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-168">
        - MatrixCoercion</span></span><br><span data-ttu-id="57f70-169">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="57f70-169">
        - Selection</span></span><br><span data-ttu-id="57f70-170">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="57f70-170">
        - Settings</span></span><br><span data-ttu-id="57f70-171">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-171">
        - TableBindings</span></span><br><span data-ttu-id="57f70-172">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-172">
        - TableCoercion</span></span><br><span data-ttu-id="57f70-173">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-173">
        - TextBindings</span></span><br><span data-ttu-id="57f70-174">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-174">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-175">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="57f70-175">Office 2019 on Windows</span></span><br><span data-ttu-id="57f70-176">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="57f70-176">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="57f70-177">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="57f70-177">- TaskPane</span></span><br><span data-ttu-id="57f70-178">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="57f70-178">
        - Content</span></span><br><span data-ttu-id="57f70-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="57f70-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="57f70-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="57f70-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="57f70-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="57f70-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="57f70-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="57f70-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="57f70-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="57f70-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="57f70-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="57f70-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="57f70-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="57f70-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="57f70-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="57f70-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="57f70-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="57f70-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="57f70-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="57f70-190">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-190">- BindingEvents</span></span><br><span data-ttu-id="57f70-191">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="57f70-191">
        - CompressedFile</span></span><br><span data-ttu-id="57f70-192">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-192">
        - DocumentEvents</span></span><br><span data-ttu-id="57f70-193">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="57f70-193">
        - File</span></span><br><span data-ttu-id="57f70-194">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-194">
        - MatrixBindings</span></span><br><span data-ttu-id="57f70-195">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-195">
        - MatrixCoercion</span></span><br><span data-ttu-id="57f70-196">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="57f70-196">
        - Selection</span></span><br><span data-ttu-id="57f70-197">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="57f70-197">
        - Settings</span></span><br><span data-ttu-id="57f70-198">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-198">
        - TableBindings</span></span><br><span data-ttu-id="57f70-199">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-199">
        - TableCoercion</span></span><br><span data-ttu-id="57f70-200">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-200">
        - TextBindings</span></span><br><span data-ttu-id="57f70-201">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-201">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-202">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="57f70-202">Office 2016 on Windows</span></span><br><span data-ttu-id="57f70-203">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="57f70-203">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="57f70-204">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="57f70-204">- TaskPane</span></span><br><span data-ttu-id="57f70-205">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="57f70-205">
        - Content</span></span></td>
    <td><span data-ttu-id="57f70-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="57f70-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="57f70-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="57f70-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="57f70-209">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-209">- BindingEvents</span></span><br><span data-ttu-id="57f70-210">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="57f70-210">
        - CompressedFile</span></span><br><span data-ttu-id="57f70-211">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-211">
        - DocumentEvents</span></span><br><span data-ttu-id="57f70-212">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="57f70-212">
        - File</span></span><br><span data-ttu-id="57f70-213">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-213">
        - MatrixBindings</span></span><br><span data-ttu-id="57f70-214">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-214">
        - MatrixCoercion</span></span><br><span data-ttu-id="57f70-215">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="57f70-215">
        - Selection</span></span><br><span data-ttu-id="57f70-216">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="57f70-216">
        - Settings</span></span><br><span data-ttu-id="57f70-217">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-217">
        - TableBindings</span></span><br><span data-ttu-id="57f70-218">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-218">
        - TableCoercion</span></span><br><span data-ttu-id="57f70-219">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-219">
        - TextBindings</span></span><br><span data-ttu-id="57f70-220">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-220">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-221">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="57f70-221">Office 2013 on Windows</span></span><br><span data-ttu-id="57f70-222">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="57f70-222">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="57f70-223">
        - TaskPane</span><span class="sxs-lookup"><span data-stu-id="57f70-223">
        - TaskPane</span></span><br><span data-ttu-id="57f70-224">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="57f70-224">
        - Content</span></span></td>
    <td>  <span data-ttu-id="57f70-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="57f70-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="57f70-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="57f70-227">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-227">
        - BindingEvents</span></span><br><span data-ttu-id="57f70-228">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-228">
        - DocumentEvents</span></span><br><span data-ttu-id="57f70-229">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="57f70-229">
        - File</span></span><br><span data-ttu-id="57f70-230">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-230">
        - MatrixBindings</span></span><br><span data-ttu-id="57f70-231">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-231">
        - MatrixCoercion</span></span><br><span data-ttu-id="57f70-232">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="57f70-232">
        - Selection</span></span><br><span data-ttu-id="57f70-233">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="57f70-233">
        - Settings</span></span><br><span data-ttu-id="57f70-234">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-234">
        - TableBindings</span></span><br><span data-ttu-id="57f70-235">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-235">
        - TableCoercion</span></span><br><span data-ttu-id="57f70-236">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-236">
        - TextBindings</span></span><br><span data-ttu-id="57f70-237">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-237">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-238">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="57f70-238">Office on iPad</span></span><br><span data-ttu-id="57f70-239">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="57f70-239">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="57f70-240">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="57f70-240">- TaskPane</span></span><br><span data-ttu-id="57f70-241">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="57f70-241">
        - Content</span></span></td>
    <td><span data-ttu-id="57f70-242">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-242">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="57f70-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="57f70-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="57f70-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="57f70-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="57f70-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="57f70-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="57f70-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="57f70-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="57f70-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="57f70-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="57f70-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="57f70-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="57f70-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="57f70-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="57f70-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="57f70-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="57f70-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="57f70-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="57f70-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="57f70-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="57f70-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="57f70-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="57f70-255">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-255">- BindingEvents</span></span><br><span data-ttu-id="57f70-256">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-256">
        - DocumentEvents</span></span><br><span data-ttu-id="57f70-257">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="57f70-257">
        - File</span></span><br><span data-ttu-id="57f70-258">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-258">
        - MatrixBindings</span></span><br><span data-ttu-id="57f70-259">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-259">
        - MatrixCoercion</span></span><br><span data-ttu-id="57f70-260">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="57f70-260">
        - Selection</span></span><br><span data-ttu-id="57f70-261">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="57f70-261">
        - Settings</span></span><br><span data-ttu-id="57f70-262">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-262">
        - TableBindings</span></span><br><span data-ttu-id="57f70-263">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-263">
        - TableCoercion</span></span><br><span data-ttu-id="57f70-264">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-264">
        - TextBindings</span></span><br><span data-ttu-id="57f70-265">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-265">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-266">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="57f70-266">Office on Mac</span></span><br><span data-ttu-id="57f70-267">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="57f70-267">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="57f70-268">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="57f70-268">- TaskPane</span></span><br><span data-ttu-id="57f70-269">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="57f70-269">
        - Content</span></span><br><span data-ttu-id="57f70-270">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="57f70-270">
        - Custom Functions</span></span><br><span data-ttu-id="57f70-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="57f70-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="57f70-272">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-272">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="57f70-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="57f70-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="57f70-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="57f70-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="57f70-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="57f70-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="57f70-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="57f70-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="57f70-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="57f70-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="57f70-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="57f70-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="57f70-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="57f70-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="57f70-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="57f70-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="57f70-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="57f70-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="57f70-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="57f70-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="57f70-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="57f70-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="57f70-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="57f70-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="57f70-286">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-286">- BindingEvents</span></span><br><span data-ttu-id="57f70-287">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="57f70-287">
        - CompressedFile</span></span><br><span data-ttu-id="57f70-288">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-288">
        - DocumentEvents</span></span><br><span data-ttu-id="57f70-289">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="57f70-289">
        - File</span></span><br><span data-ttu-id="57f70-290">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-290">
        - MatrixBindings</span></span><br><span data-ttu-id="57f70-291">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-291">
        - MatrixCoercion</span></span><br><span data-ttu-id="57f70-292">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="57f70-292">
        - PdfFile</span></span><br><span data-ttu-id="57f70-293">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="57f70-293">
        - Selection</span></span><br><span data-ttu-id="57f70-294">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="57f70-294">
        - Settings</span></span><br><span data-ttu-id="57f70-295">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-295">
        - TableBindings</span></span><br><span data-ttu-id="57f70-296">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-296">
        - TableCoercion</span></span><br><span data-ttu-id="57f70-297">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-297">
        - TextBindings</span></span><br><span data-ttu-id="57f70-298">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-298">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-299">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="57f70-299">Office 2019 on Mac</span></span><br><span data-ttu-id="57f70-300">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="57f70-300">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="57f70-301">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="57f70-301">- TaskPane</span></span><br><span data-ttu-id="57f70-302">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="57f70-302">
        - Content</span></span><br><span data-ttu-id="57f70-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="57f70-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="57f70-304">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-304">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="57f70-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="57f70-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="57f70-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="57f70-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="57f70-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="57f70-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="57f70-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="57f70-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="57f70-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="57f70-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="57f70-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="57f70-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="57f70-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="57f70-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="57f70-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="57f70-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="57f70-314">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-314">- BindingEvents</span></span><br><span data-ttu-id="57f70-315">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="57f70-315">
        - CompressedFile</span></span><br><span data-ttu-id="57f70-316">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-316">
        - DocumentEvents</span></span><br><span data-ttu-id="57f70-317">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="57f70-317">
        - File</span></span><br><span data-ttu-id="57f70-318">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-318">
        - MatrixBindings</span></span><br><span data-ttu-id="57f70-319">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-319">
        - MatrixCoercion</span></span><br><span data-ttu-id="57f70-320">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="57f70-320">
        - PdfFile</span></span><br><span data-ttu-id="57f70-321">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="57f70-321">
        - Selection</span></span><br><span data-ttu-id="57f70-322">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="57f70-322">
        - Settings</span></span><br><span data-ttu-id="57f70-323">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-323">
        - TableBindings</span></span><br><span data-ttu-id="57f70-324">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-324">
        - TableCoercion</span></span><br><span data-ttu-id="57f70-325">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-325">
        - TextBindings</span></span><br><span data-ttu-id="57f70-326">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-326">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-327">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="57f70-327">Office 2016 on Mac</span></span><br><span data-ttu-id="57f70-328">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="57f70-328">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="57f70-329">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="57f70-329">- TaskPane</span></span><br><span data-ttu-id="57f70-330">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="57f70-330">
        - Content</span></span></td>
    <td><span data-ttu-id="57f70-331">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-331">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="57f70-332">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="57f70-332">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="57f70-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="57f70-334">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-334">- BindingEvents</span></span><br><span data-ttu-id="57f70-335">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="57f70-335">
        - CompressedFile</span></span><br><span data-ttu-id="57f70-336">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-336">
        - DocumentEvents</span></span><br><span data-ttu-id="57f70-337">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="57f70-337">
        - File</span></span><br><span data-ttu-id="57f70-338">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-338">
        - MatrixBindings</span></span><br><span data-ttu-id="57f70-339">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-339">
        - MatrixCoercion</span></span><br><span data-ttu-id="57f70-340">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="57f70-340">
        - PdfFile</span></span><br><span data-ttu-id="57f70-341">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="57f70-341">
        - Selection</span></span><br><span data-ttu-id="57f70-342">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="57f70-342">
        - Settings</span></span><br><span data-ttu-id="57f70-343">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-343">
        - TableBindings</span></span><br><span data-ttu-id="57f70-344">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-344">
        - TableCoercion</span></span><br><span data-ttu-id="57f70-345">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-345">
        - TextBindings</span></span><br><span data-ttu-id="57f70-346">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-346">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="57f70-347">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="57f70-347">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="57f70-348">Funções personalizadas (somente Excel)</span><span class="sxs-lookup"><span data-stu-id="57f70-348">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="57f70-349">Plataforma</span><span class="sxs-lookup"><span data-stu-id="57f70-349">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="57f70-350">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="57f70-350">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="57f70-351">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="57f70-351">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="57f70-352"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="57f70-352"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-353">Office na Web</span><span class="sxs-lookup"><span data-stu-id="57f70-353">Office on the web</span></span></td>
    <td><span data-ttu-id="57f70-354">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="57f70-354">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="57f70-355">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-355">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-356">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="57f70-356">Office on Windows</span></span><br><span data-ttu-id="57f70-357">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="57f70-357">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="57f70-358">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="57f70-358">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="57f70-359">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-359">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-360">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="57f70-360">Office on Mac</span></span><br><span data-ttu-id="57f70-361">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="57f70-361">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="57f70-362">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="57f70-362">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="57f70-363">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-363">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="57f70-364">Outlook</span><span class="sxs-lookup"><span data-stu-id="57f70-364">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="57f70-365">Plataforma</span><span class="sxs-lookup"><span data-stu-id="57f70-365">Platform</span></span></th>
    <th><span data-ttu-id="57f70-366">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="57f70-366">Extension points</span></span></th>
    <th><span data-ttu-id="57f70-367">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="57f70-367">API requirement sets</span></span></th>
    <th><span data-ttu-id="57f70-368"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="57f70-368"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-369">Office na Web</span><span class="sxs-lookup"><span data-stu-id="57f70-369">Office on the web</span></span><br><span data-ttu-id="57f70-370">(moderno)</span><span class="sxs-lookup"><span data-stu-id="57f70-370">(modern)</span></span></td>
    <td> <span data-ttu-id="57f70-371">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="57f70-371">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="57f70-372">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="57f70-372">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="57f70-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="57f70-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="57f70-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="57f70-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="57f70-375">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="57f70-375">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="57f70-376">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-376">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="57f70-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="57f70-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="57f70-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="57f70-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="57f70-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="57f70-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="57f70-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="57f70-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="57f70-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="57f70-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="57f70-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="57f70-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="57f70-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a></span><span class="sxs-lookup"><span data-stu-id="57f70-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="57f70-384">Não disponível</span><span class="sxs-lookup"><span data-stu-id="57f70-384">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-385">Office na Web</span><span class="sxs-lookup"><span data-stu-id="57f70-385">Office on the web</span></span><br><span data-ttu-id="57f70-386">(clássico)</span><span class="sxs-lookup"><span data-stu-id="57f70-386">(classic)</span></span></td>
    <td> <span data-ttu-id="57f70-387">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="57f70-387">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="57f70-388">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="57f70-388">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="57f70-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="57f70-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="57f70-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="57f70-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="57f70-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="57f70-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="57f70-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="57f70-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="57f70-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="57f70-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="57f70-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="57f70-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="57f70-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="57f70-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="57f70-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="57f70-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="57f70-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="57f70-398">Não disponível</span><span class="sxs-lookup"><span data-stu-id="57f70-398">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-399">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="57f70-399">Office on Windows</span></span><br><span data-ttu-id="57f70-400">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="57f70-400">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="57f70-401">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="57f70-401">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="57f70-402">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="57f70-402">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="57f70-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="57f70-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="57f70-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="57f70-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="57f70-405">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="57f70-405">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="57f70-406">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Módulos</a></span><span class="sxs-lookup"><span data-stu-id="57f70-406">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="57f70-407">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-407">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="57f70-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="57f70-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="57f70-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="57f70-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="57f70-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="57f70-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="57f70-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="57f70-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="57f70-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="57f70-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="57f70-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="57f70-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="57f70-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a></span><span class="sxs-lookup"><span data-stu-id="57f70-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="57f70-415">Não disponível</span><span class="sxs-lookup"><span data-stu-id="57f70-415">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-416">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="57f70-416">Office 2019 on Windows</span></span><br><span data-ttu-id="57f70-417">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="57f70-417">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="57f70-418">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="57f70-418">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="57f70-419">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="57f70-419">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="57f70-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="57f70-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="57f70-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="57f70-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="57f70-422">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="57f70-422">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="57f70-423">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Módulos</a></span><span class="sxs-lookup"><span data-stu-id="57f70-423">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="57f70-424">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-424">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="57f70-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="57f70-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="57f70-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="57f70-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="57f70-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="57f70-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="57f70-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="57f70-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="57f70-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="57f70-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="57f70-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="57f70-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="57f70-431">Não disponível</span><span class="sxs-lookup"><span data-stu-id="57f70-431">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-432">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="57f70-432">Office 2016 on Windows</span></span><br><span data-ttu-id="57f70-433">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="57f70-433">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="57f70-434">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="57f70-434">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="57f70-435">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="57f70-435">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="57f70-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="57f70-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="57f70-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="57f70-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="57f70-438">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="57f70-438">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="57f70-439">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Módulos</a></span><span class="sxs-lookup"><span data-stu-id="57f70-439">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="57f70-440">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-440">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="57f70-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="57f70-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="57f70-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="57f70-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="57f70-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="57f70-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="57f70-444">Não disponível</span><span class="sxs-lookup"><span data-stu-id="57f70-444">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-445">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="57f70-445">Office 2013 on Windows</span></span><br><span data-ttu-id="57f70-446">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="57f70-446">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="57f70-447">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="57f70-447">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="57f70-448">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="57f70-448">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="57f70-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="57f70-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="57f70-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="57f70-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br>
    <td> <span data-ttu-id="57f70-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="57f70-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="57f70-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="57f70-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="57f70-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="57f70-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="57f70-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="57f70-455">Não disponível</span><span class="sxs-lookup"><span data-stu-id="57f70-455">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-456">Office no iOS</span><span class="sxs-lookup"><span data-stu-id="57f70-456">Office on iOS</span></span><br><span data-ttu-id="57f70-457">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="57f70-457">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="57f70-458">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="57f70-458">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="57f70-459">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="57f70-459">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="57f70-460">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-460">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="57f70-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="57f70-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="57f70-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="57f70-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="57f70-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="57f70-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="57f70-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="57f70-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="57f70-465">Não disponível</span><span class="sxs-lookup"><span data-stu-id="57f70-465">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-466">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="57f70-466">Office on Mac</span></span><br><span data-ttu-id="57f70-467">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="57f70-467">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="57f70-468">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="57f70-468">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="57f70-469">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="57f70-469">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="57f70-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="57f70-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="57f70-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="57f70-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="57f70-472">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="57f70-472">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="57f70-473">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-473">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="57f70-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="57f70-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="57f70-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="57f70-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="57f70-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="57f70-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="57f70-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="57f70-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="57f70-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="57f70-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="57f70-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="57f70-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="57f70-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a></span><span class="sxs-lookup"><span data-stu-id="57f70-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="57f70-481">Não disponível</span><span class="sxs-lookup"><span data-stu-id="57f70-481">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-482">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="57f70-482">Office 2019 on Mac</span></span><br><span data-ttu-id="57f70-483">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="57f70-483">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="57f70-484">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="57f70-484">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="57f70-485">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="57f70-485">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="57f70-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="57f70-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="57f70-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="57f70-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="57f70-488">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="57f70-488">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="57f70-489">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-489">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="57f70-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="57f70-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="57f70-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="57f70-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="57f70-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="57f70-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="57f70-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="57f70-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="57f70-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="57f70-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="57f70-495">Não disponível</span><span class="sxs-lookup"><span data-stu-id="57f70-495">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-496">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="57f70-496">Office 2016 on Mac</span></span><br><span data-ttu-id="57f70-497">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="57f70-497">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="57f70-498">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="57f70-498">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="57f70-499">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="57f70-499">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="57f70-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="57f70-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="57f70-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="57f70-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="57f70-502">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="57f70-502">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="57f70-503">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-503">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="57f70-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="57f70-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="57f70-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="57f70-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="57f70-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="57f70-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="57f70-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="57f70-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="57f70-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="57f70-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="57f70-509">Não disponível</span><span class="sxs-lookup"><span data-stu-id="57f70-509">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-510">Outlook no Android</span><span class="sxs-lookup"><span data-stu-id="57f70-510">Office on Android</span></span><br><span data-ttu-id="57f70-511">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="57f70-511">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="57f70-512">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="57f70-512">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="57f70-513">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Organizador de compromissos (Redigir): reunião on-line (visualização)</span><span class="sxs-lookup"><span data-stu-id="57f70-513">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Appointment Organizer (Compose): online meeting</a> (preview)</span></span><br><span data-ttu-id="57f70-514">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="57f70-514">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="57f70-515">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-515">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="57f70-516">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="57f70-516">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="57f70-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="57f70-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="57f70-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="57f70-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="57f70-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="57f70-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="57f70-520">Não disponível</span><span class="sxs-lookup"><span data-stu-id="57f70-520">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="57f70-521">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="57f70-521">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="57f70-522">O suporte ao cliente para um conjunto de requisitos pode ser restringido pelo suporte do servidor Exchange.</span><span class="sxs-lookup"><span data-stu-id="57f70-522">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="57f70-523">Consulte [Conjuntos de requisitos da API JavaScript do Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) para obter detalhes sobre o intervalo de conjuntos de requisitos suportado pelo servidor Exchange e pelos clientes Outlook.</span><span class="sxs-lookup"><span data-stu-id="57f70-523">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="57f70-524">Word</span><span class="sxs-lookup"><span data-stu-id="57f70-524">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="57f70-525">Plataforma</span><span class="sxs-lookup"><span data-stu-id="57f70-525">Platform</span></span></th>
    <th><span data-ttu-id="57f70-526">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="57f70-526">Extension points</span></span></th>
    <th><span data-ttu-id="57f70-527">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="57f70-527">API requirement sets</span></span></th>
    <th><span data-ttu-id="57f70-528"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="57f70-528"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-529">Office na Web</span><span class="sxs-lookup"><span data-stu-id="57f70-529">Office on the web</span></span></td>
    <td> <span data-ttu-id="57f70-530">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="57f70-530">- TaskPane</span></span><br><span data-ttu-id="57f70-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="57f70-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="57f70-532">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-532">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="57f70-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="57f70-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="57f70-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="57f70-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="57f70-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="57f70-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="57f70-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="57f70-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="57f70-538">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-538">- BindingEvents</span></span><br><span data-ttu-id="57f70-539">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="57f70-539">
         - CustomXmlParts</span></span><br><span data-ttu-id="57f70-540">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-540">
         - DocumentEvents</span></span><br><span data-ttu-id="57f70-541">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="57f70-541">
         - File</span></span><br><span data-ttu-id="57f70-542">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-542">
         - HtmlCoercion</span></span><br><span data-ttu-id="57f70-543">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-543">
         - MatrixBindings</span></span><br><span data-ttu-id="57f70-544">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-544">
         - MatrixCoercion</span></span><br><span data-ttu-id="57f70-545">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-545">
         - OoxmlCoercion</span></span><br><span data-ttu-id="57f70-546">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="57f70-546">
         - PdfFile</span></span><br><span data-ttu-id="57f70-547">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="57f70-547">
         - Selection</span></span><br><span data-ttu-id="57f70-548">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="57f70-548">
         - Settings</span></span><br><span data-ttu-id="57f70-549">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-549">
         - TableBindings</span></span><br><span data-ttu-id="57f70-550">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-550">
         - TableCoercion</span></span><br><span data-ttu-id="57f70-551">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-551">
         - TextBindings</span></span><br><span data-ttu-id="57f70-552">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-552">
         - TextCoercion</span></span><br><span data-ttu-id="57f70-553">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="57f70-553">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-554">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="57f70-554">Office on Windows</span></span><br><span data-ttu-id="57f70-555">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="57f70-555">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="57f70-556">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="57f70-556">- TaskPane</span></span><br><span data-ttu-id="57f70-557">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="57f70-557">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="57f70-558">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-558">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="57f70-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="57f70-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="57f70-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="57f70-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="57f70-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="57f70-562">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-562">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="57f70-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="57f70-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="57f70-564">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-564">- BindingEvents</span></span><br><span data-ttu-id="57f70-565">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="57f70-565">
         - CompressedFile</span></span><br><span data-ttu-id="57f70-566">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="57f70-566">
         - CustomXmlParts</span></span><br><span data-ttu-id="57f70-567">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-567">
         - DocumentEvents</span></span><br><span data-ttu-id="57f70-568">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="57f70-568">
         - File</span></span><br><span data-ttu-id="57f70-569">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-569">
         - HtmlCoercion</span></span><br><span data-ttu-id="57f70-570">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-570">
         - MatrixBindings</span></span><br><span data-ttu-id="57f70-571">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-571">
         - MatrixCoercion</span></span><br><span data-ttu-id="57f70-572">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-572">
         - OoxmlCoercion</span></span><br><span data-ttu-id="57f70-573">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="57f70-573">
         - PdfFile</span></span><br><span data-ttu-id="57f70-574">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="57f70-574">
         - Selection</span></span><br><span data-ttu-id="57f70-575">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="57f70-575">
         - Settings</span></span><br><span data-ttu-id="57f70-576">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-576">
         - TableBindings</span></span><br><span data-ttu-id="57f70-577">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-577">
         - TableCoercion</span></span><br><span data-ttu-id="57f70-578">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-578">
         - TextBindings</span></span><br><span data-ttu-id="57f70-579">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-579">
         - TextCoercion</span></span><br><span data-ttu-id="57f70-580">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="57f70-580">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-581">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="57f70-581">Office 2019 on Windows</span></span><br><span data-ttu-id="57f70-582">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="57f70-582">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="57f70-583">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="57f70-583">- TaskPane</span></span><br><span data-ttu-id="57f70-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="57f70-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="57f70-585">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-585">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="57f70-586">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="57f70-586">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="57f70-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="57f70-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="57f70-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="57f70-589">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-589">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="57f70-590">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-590">- BindingEvents</span></span><br><span data-ttu-id="57f70-591">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="57f70-591">
         - CompressedFile</span></span><br><span data-ttu-id="57f70-592">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="57f70-592">
         - CustomXmlParts</span></span><br><span data-ttu-id="57f70-593">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-593">
         - DocumentEvents</span></span><br><span data-ttu-id="57f70-594">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="57f70-594">
         - File</span></span><br><span data-ttu-id="57f70-595">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-595">
         - HtmlCoercion</span></span><br><span data-ttu-id="57f70-596">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-596">
         - MatrixBindings</span></span><br><span data-ttu-id="57f70-597">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-597">
         - MatrixCoercion</span></span><br><span data-ttu-id="57f70-598">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-598">
         - OoxmlCoercion</span></span><br><span data-ttu-id="57f70-599">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="57f70-599">
         - PdfFile</span></span><br><span data-ttu-id="57f70-600">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="57f70-600">
         - Selection</span></span><br><span data-ttu-id="57f70-601">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="57f70-601">
         - Settings</span></span><br><span data-ttu-id="57f70-602">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-602">
         - TableBindings</span></span><br><span data-ttu-id="57f70-603">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-603">
         - TableCoercion</span></span><br><span data-ttu-id="57f70-604">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-604">
         - TextBindings</span></span><br><span data-ttu-id="57f70-605">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-605">
         - TextCoercion</span></span><br><span data-ttu-id="57f70-606">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="57f70-606">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-607">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="57f70-607">Office 2016 on Windows</span></span><br><span data-ttu-id="57f70-608">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="57f70-608">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="57f70-609">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="57f70-609">- TaskPane</span></span></td>
    <td> <span data-ttu-id="57f70-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="57f70-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="57f70-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="57f70-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="57f70-613">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-613">- BindingEvents</span></span><br><span data-ttu-id="57f70-614">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="57f70-614">
         - CompressedFile</span></span><br><span data-ttu-id="57f70-615">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="57f70-615">
         - CustomXmlParts</span></span><br><span data-ttu-id="57f70-616">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-616">
         - DocumentEvents</span></span><br><span data-ttu-id="57f70-617">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="57f70-617">
         - File</span></span><br><span data-ttu-id="57f70-618">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-618">
         - HtmlCoercion</span></span><br><span data-ttu-id="57f70-619">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-619">
         - MatrixBindings</span></span><br><span data-ttu-id="57f70-620">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-620">
         - MatrixCoercion</span></span><br><span data-ttu-id="57f70-621">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-621">
         - OoxmlCoercion</span></span><br><span data-ttu-id="57f70-622">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="57f70-622">
         - PdfFile</span></span><br><span data-ttu-id="57f70-623">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="57f70-623">
         - Selection</span></span><br><span data-ttu-id="57f70-624">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="57f70-624">
         - Settings</span></span><br><span data-ttu-id="57f70-625">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-625">
         - TableBindings</span></span><br><span data-ttu-id="57f70-626">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-626">
         - TableCoercion</span></span><br><span data-ttu-id="57f70-627">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-627">
         - TextBindings</span></span><br><span data-ttu-id="57f70-628">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-628">
         - TextCoercion</span></span><br><span data-ttu-id="57f70-629">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="57f70-629">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-630">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="57f70-630">Office 2013 on Windows</span></span><br><span data-ttu-id="57f70-631">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="57f70-631">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="57f70-632">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="57f70-632">- TaskPane</span></span></td>
    <td> <span data-ttu-id="57f70-633">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="57f70-633">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="57f70-634">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-634">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="57f70-635">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-635">- BindingEvents</span></span><br><span data-ttu-id="57f70-636">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="57f70-636">
         - CompressedFile</span></span><br><span data-ttu-id="57f70-637">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="57f70-637">
         - CustomXmlParts</span></span><br><span data-ttu-id="57f70-638">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-638">
         - DocumentEvents</span></span><br><span data-ttu-id="57f70-639">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="57f70-639">
         - File</span></span><br><span data-ttu-id="57f70-640">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-640">
         - HtmlCoercion</span></span><br><span data-ttu-id="57f70-641">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-641">
         - MatrixBindings</span></span><br><span data-ttu-id="57f70-642">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-642">
         - MatrixCoercion</span></span><br><span data-ttu-id="57f70-643">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-643">
         - OoxmlCoercion</span></span><br><span data-ttu-id="57f70-644">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="57f70-644">
         - PdfFile</span></span><br><span data-ttu-id="57f70-645">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="57f70-645">
         - Selection</span></span><br><span data-ttu-id="57f70-646">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="57f70-646">
         - Settings</span></span><br><span data-ttu-id="57f70-647">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-647">
         - TableBindings</span></span><br><span data-ttu-id="57f70-648">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-648">
         - TableCoercion</span></span><br><span data-ttu-id="57f70-649">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-649">
         - TextBindings</span></span><br><span data-ttu-id="57f70-650">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-650">
         - TextCoercion</span></span><br><span data-ttu-id="57f70-651">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="57f70-651">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-652">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="57f70-652">Office on iPad</span></span><br><span data-ttu-id="57f70-653">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="57f70-653">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="57f70-654">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="57f70-654">- TaskPane</span></span></td>
    <td> <span data-ttu-id="57f70-655">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-655">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="57f70-656">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="57f70-656">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="57f70-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="57f70-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="57f70-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="57f70-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="57f70-660">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-660">- BindingEvents</span></span><br><span data-ttu-id="57f70-661">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="57f70-661">
         - CompressedFile</span></span><br><span data-ttu-id="57f70-662">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="57f70-662">
         - CustomXmlParts</span></span><br><span data-ttu-id="57f70-663">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-663">
         - DocumentEvents</span></span><br><span data-ttu-id="57f70-664">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="57f70-664">
         - File</span></span><br><span data-ttu-id="57f70-665">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-665">
         - HtmlCoercion</span></span><br><span data-ttu-id="57f70-666">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-666">
         - MatrixBindings</span></span><br><span data-ttu-id="57f70-667">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-667">
         - MatrixCoercion</span></span><br><span data-ttu-id="57f70-668">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-668">
         - OoxmlCoercion</span></span><br><span data-ttu-id="57f70-669">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="57f70-669">
         - PdfFile</span></span><br><span data-ttu-id="57f70-670">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="57f70-670">
         - Selection</span></span><br><span data-ttu-id="57f70-671">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="57f70-671">
         - Settings</span></span><br><span data-ttu-id="57f70-672">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-672">
         - TableBindings</span></span><br><span data-ttu-id="57f70-673">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-673">
         - TableCoercion</span></span><br><span data-ttu-id="57f70-674">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-674">
         - TextBindings</span></span><br><span data-ttu-id="57f70-675">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-675">
         - TextCoercion</span></span><br><span data-ttu-id="57f70-676">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="57f70-676">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-677">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="57f70-677">Office on Mac</span></span><br><span data-ttu-id="57f70-678">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="57f70-678">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="57f70-679">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="57f70-679">- TaskPane</span></span><br><span data-ttu-id="57f70-680">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="57f70-680">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="57f70-681">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-681">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="57f70-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="57f70-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="57f70-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="57f70-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="57f70-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="57f70-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="57f70-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="57f70-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="57f70-687">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-687">- BindingEvents</span></span><br><span data-ttu-id="57f70-688">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="57f70-688">
         - CompressedFile</span></span><br><span data-ttu-id="57f70-689">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="57f70-689">
         - CustomXmlParts</span></span><br><span data-ttu-id="57f70-690">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-690">
         - DocumentEvents</span></span><br><span data-ttu-id="57f70-691">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="57f70-691">
         - File</span></span><br><span data-ttu-id="57f70-692">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-692">
         - HtmlCoercion</span></span><br><span data-ttu-id="57f70-693">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-693">
         - MatrixBindings</span></span><br><span data-ttu-id="57f70-694">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-694">
         - MatrixCoercion</span></span><br><span data-ttu-id="57f70-695">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-695">
         - OoxmlCoercion</span></span><br><span data-ttu-id="57f70-696">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="57f70-696">
         - PdfFile</span></span><br><span data-ttu-id="57f70-697">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="57f70-697">
         - Selection</span></span><br><span data-ttu-id="57f70-698">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="57f70-698">
         - Settings</span></span><br><span data-ttu-id="57f70-699">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-699">
         - TableBindings</span></span><br><span data-ttu-id="57f70-700">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-700">
         - TableCoercion</span></span><br><span data-ttu-id="57f70-701">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-701">
         - TextBindings</span></span><br><span data-ttu-id="57f70-702">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-702">
         - TextCoercion</span></span><br><span data-ttu-id="57f70-703">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="57f70-703">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-704">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="57f70-704">Office 2019 on Mac</span></span><br><span data-ttu-id="57f70-705">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="57f70-705">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="57f70-706">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="57f70-706">- TaskPane</span></span><br><span data-ttu-id="57f70-707">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="57f70-707">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="57f70-708">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-708">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="57f70-709">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="57f70-709">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="57f70-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="57f70-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="57f70-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="57f70-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="57f70-713">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-713">- BindingEvents</span></span><br><span data-ttu-id="57f70-714">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="57f70-714">
         - CompressedFile</span></span><br><span data-ttu-id="57f70-715">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="57f70-715">
         - CustomXmlParts</span></span><br><span data-ttu-id="57f70-716">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-716">
         - DocumentEvents</span></span><br><span data-ttu-id="57f70-717">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="57f70-717">
         - File</span></span><br><span data-ttu-id="57f70-718">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-718">
         - HtmlCoercion</span></span><br><span data-ttu-id="57f70-719">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-719">
         - MatrixBindings</span></span><br><span data-ttu-id="57f70-720">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-720">
         - MatrixCoercion</span></span><br><span data-ttu-id="57f70-721">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-721">
         - OoxmlCoercion</span></span><br><span data-ttu-id="57f70-722">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="57f70-722">
         - PdfFile</span></span><br><span data-ttu-id="57f70-723">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="57f70-723">
         - Selection</span></span><br><span data-ttu-id="57f70-724">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="57f70-724">
         - Settings</span></span><br><span data-ttu-id="57f70-725">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-725">
         - TableBindings</span></span><br><span data-ttu-id="57f70-726">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-726">
         - TableCoercion</span></span><br><span data-ttu-id="57f70-727">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-727">
         - TextBindings</span></span><br><span data-ttu-id="57f70-728">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-728">
         - TextCoercion</span></span><br><span data-ttu-id="57f70-729">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="57f70-729">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-730">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="57f70-730">Office 2016 on Mac</span></span><br><span data-ttu-id="57f70-731">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="57f70-731">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="57f70-732">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="57f70-732">- TaskPane</span></span></td>
    <td> <span data-ttu-id="57f70-733">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-733">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="57f70-734">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="57f70-734">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="57f70-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="57f70-736">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-736">- BindingEvents</span></span><br><span data-ttu-id="57f70-737">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="57f70-737">
         - CompressedFile</span></span><br><span data-ttu-id="57f70-738">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="57f70-738">
         - CustomXmlParts</span></span><br><span data-ttu-id="57f70-739">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-739">
         - DocumentEvents</span></span><br><span data-ttu-id="57f70-740">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="57f70-740">
         - File</span></span><br><span data-ttu-id="57f70-741">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-741">
         - HtmlCoercion</span></span><br><span data-ttu-id="57f70-742">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-742">
         - MatrixBindings</span></span><br><span data-ttu-id="57f70-743">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-743">
         - MatrixCoercion</span></span><br><span data-ttu-id="57f70-744">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-744">
         - OoxmlCoercion</span></span><br><span data-ttu-id="57f70-745">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="57f70-745">
         - PdfFile</span></span><br><span data-ttu-id="57f70-746">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="57f70-746">
         - Selection</span></span><br><span data-ttu-id="57f70-747">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="57f70-747">
         - Settings</span></span><br><span data-ttu-id="57f70-748">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-748">
         - TableBindings</span></span><br><span data-ttu-id="57f70-749">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-749">
         - TableCoercion</span></span><br><span data-ttu-id="57f70-750">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="57f70-750">
         - TextBindings</span></span><br><span data-ttu-id="57f70-751">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-751">
         - TextCoercion</span></span><br><span data-ttu-id="57f70-752">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="57f70-752">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="57f70-753">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="57f70-753">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="57f70-754">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="57f70-754">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="57f70-755">Plataforma</span><span class="sxs-lookup"><span data-stu-id="57f70-755">Platform</span></span></th>
    <th><span data-ttu-id="57f70-756">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="57f70-756">Extension points</span></span></th>
    <th><span data-ttu-id="57f70-757">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="57f70-757">API requirement sets</span></span></th>
    <th><span data-ttu-id="57f70-758"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="57f70-758"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-759">Office na Web</span><span class="sxs-lookup"><span data-stu-id="57f70-759">Office on the web</span></span></td>
    <td> <span data-ttu-id="57f70-760">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="57f70-760">- Content</span></span><br><span data-ttu-id="57f70-761">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="57f70-761">
         - TaskPane</span></span><br><span data-ttu-id="57f70-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="57f70-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="57f70-763">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-763">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="57f70-764">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-764">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="57f70-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="57f70-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="57f70-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="57f70-767">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="57f70-767">- ActiveView</span></span><br><span data-ttu-id="57f70-768">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="57f70-768">
         - CompressedFile</span></span><br><span data-ttu-id="57f70-769">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-769">
         - DocumentEvents</span></span><br><span data-ttu-id="57f70-770">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="57f70-770">
         - File</span></span><br><span data-ttu-id="57f70-771">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="57f70-771">
         - PdfFile</span></span><br><span data-ttu-id="57f70-772">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="57f70-772">
         - Selection</span></span><br><span data-ttu-id="57f70-773">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="57f70-773">
         - Settings</span></span><br><span data-ttu-id="57f70-774">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-774">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-775">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="57f70-775">Office on Windows</span></span><br><span data-ttu-id="57f70-776">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="57f70-776">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="57f70-777">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="57f70-777">- Content</span></span><br><span data-ttu-id="57f70-778">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="57f70-778">
         - TaskPane</span></span><br><span data-ttu-id="57f70-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="57f70-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="57f70-780">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-780">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="57f70-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="57f70-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="57f70-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="57f70-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="57f70-784">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="57f70-784">- ActiveView</span></span><br><span data-ttu-id="57f70-785">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="57f70-785">
         - CompressedFile</span></span><br><span data-ttu-id="57f70-786">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-786">
         - DocumentEvents</span></span><br><span data-ttu-id="57f70-787">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="57f70-787">
         - File</span></span><br><span data-ttu-id="57f70-788">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="57f70-788">
         - PdfFile</span></span><br><span data-ttu-id="57f70-789">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="57f70-789">
         - Selection</span></span><br><span data-ttu-id="57f70-790">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="57f70-790">
         - Settings</span></span><br><span data-ttu-id="57f70-791">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-791">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-792">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="57f70-792">Office 2019 on Windows</span></span><br><span data-ttu-id="57f70-793">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="57f70-793">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="57f70-794">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="57f70-794">- Content</span></span><br><span data-ttu-id="57f70-795">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="57f70-795">
         - TaskPane</span></span><br><span data-ttu-id="57f70-796">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="57f70-796">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="57f70-797">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-797">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="57f70-798">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-798">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="57f70-799">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="57f70-799">- ActiveView</span></span><br><span data-ttu-id="57f70-800">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="57f70-800">
         - CompressedFile</span></span><br><span data-ttu-id="57f70-801">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-801">
         - DocumentEvents</span></span><br><span data-ttu-id="57f70-802">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="57f70-802">
         - File</span></span><br><span data-ttu-id="57f70-803">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="57f70-803">
         - PdfFile</span></span><br><span data-ttu-id="57f70-804">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="57f70-804">
         - Selection</span></span><br><span data-ttu-id="57f70-805">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="57f70-805">
         - Settings</span></span><br><span data-ttu-id="57f70-806">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-806">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-807">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="57f70-807">Office 2016 on Windows</span></span><br><span data-ttu-id="57f70-808">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="57f70-808">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="57f70-809">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="57f70-809">- Content</span></span><br><span data-ttu-id="57f70-810">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="57f70-810">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="57f70-811">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="57f70-811">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="57f70-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="57f70-813">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="57f70-813">- ActiveView</span></span><br><span data-ttu-id="57f70-814">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="57f70-814">
         - CompressedFile</span></span><br><span data-ttu-id="57f70-815">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-815">
         - DocumentEvents</span></span><br><span data-ttu-id="57f70-816">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="57f70-816">
         - File</span></span><br><span data-ttu-id="57f70-817">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="57f70-817">
         - PdfFile</span></span><br><span data-ttu-id="57f70-818">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="57f70-818">
         - Selection</span></span><br><span data-ttu-id="57f70-819">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="57f70-819">
         - Settings</span></span><br><span data-ttu-id="57f70-820">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-820">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-821">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="57f70-821">Office 2013 on Windows</span></span><br><span data-ttu-id="57f70-822">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="57f70-822">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="57f70-823">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="57f70-823">- Content</span></span><br><span data-ttu-id="57f70-824">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="57f70-824">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="57f70-825">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="57f70-825">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="57f70-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="57f70-827">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="57f70-827">- ActiveView</span></span><br><span data-ttu-id="57f70-828">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="57f70-828">
         - CompressedFile</span></span><br><span data-ttu-id="57f70-829">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-829">
         - DocumentEvents</span></span><br><span data-ttu-id="57f70-830">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="57f70-830">
         - File</span></span><br><span data-ttu-id="57f70-831">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="57f70-831">
         - PdfFile</span></span><br><span data-ttu-id="57f70-832">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="57f70-832">
         - Selection</span></span><br><span data-ttu-id="57f70-833">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="57f70-833">
         - Settings</span></span><br><span data-ttu-id="57f70-834">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-834">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-835">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="57f70-835">Office on iPad</span></span><br><span data-ttu-id="57f70-836">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="57f70-836">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="57f70-837">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="57f70-837">- Content</span></span><br><span data-ttu-id="57f70-838">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="57f70-838">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="57f70-839">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-839">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="57f70-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="57f70-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="57f70-842">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="57f70-842">- ActiveView</span></span><br><span data-ttu-id="57f70-843">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="57f70-843">
         - CompressedFile</span></span><br><span data-ttu-id="57f70-844">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-844">
         - DocumentEvents</span></span><br><span data-ttu-id="57f70-845">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="57f70-845">
         - File</span></span><br><span data-ttu-id="57f70-846">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="57f70-846">
         - PdfFile</span></span><br><span data-ttu-id="57f70-847">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="57f70-847">
         - Selection</span></span><br><span data-ttu-id="57f70-848">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="57f70-848">
         - Settings</span></span><br><span data-ttu-id="57f70-849">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-849">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-850">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="57f70-850">Office on Mac</span></span><br><span data-ttu-id="57f70-851">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="57f70-851">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="57f70-852">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="57f70-852">- Content</span></span><br><span data-ttu-id="57f70-853">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="57f70-853">
         - TaskPane</span></span><br><span data-ttu-id="57f70-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="57f70-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="57f70-855">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-855">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="57f70-856">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-856">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="57f70-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="57f70-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="57f70-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="57f70-859">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="57f70-859">- ActiveView</span></span><br><span data-ttu-id="57f70-860">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="57f70-860">
         - CompressedFile</span></span><br><span data-ttu-id="57f70-861">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-861">
         - DocumentEvents</span></span><br><span data-ttu-id="57f70-862">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="57f70-862">
         - File</span></span><br><span data-ttu-id="57f70-863">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="57f70-863">
         - PdfFile</span></span><br><span data-ttu-id="57f70-864">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="57f70-864">
         - Selection</span></span><br><span data-ttu-id="57f70-865">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="57f70-865">
         - Settings</span></span><br><span data-ttu-id="57f70-866">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-866">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-867">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="57f70-867">Office 2019 on Mac</span></span><br><span data-ttu-id="57f70-868">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="57f70-868">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="57f70-869">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="57f70-869">- Content</span></span><br><span data-ttu-id="57f70-870">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="57f70-870">
         - TaskPane</span></span><br><span data-ttu-id="57f70-871">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="57f70-871">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="57f70-872">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-872">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="57f70-873">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-873">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="57f70-874">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="57f70-874">- ActiveView</span></span><br><span data-ttu-id="57f70-875">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="57f70-875">
         - CompressedFile</span></span><br><span data-ttu-id="57f70-876">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-876">
         - DocumentEvents</span></span><br><span data-ttu-id="57f70-877">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="57f70-877">
         - File</span></span><br><span data-ttu-id="57f70-878">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="57f70-878">
         - PdfFile</span></span><br><span data-ttu-id="57f70-879">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="57f70-879">
         - Selection</span></span><br><span data-ttu-id="57f70-880">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="57f70-880">
         - Settings</span></span><br><span data-ttu-id="57f70-881">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-881">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-882">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="57f70-882">Office 2016 on Mac</span></span><br><span data-ttu-id="57f70-883">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="57f70-883">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="57f70-884">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="57f70-884">- Content</span></span><br><span data-ttu-id="57f70-885">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="57f70-885">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="57f70-886">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="57f70-886">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="57f70-887">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-887">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="57f70-888">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="57f70-888">- ActiveView</span></span><br><span data-ttu-id="57f70-889">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="57f70-889">
         - CompressedFile</span></span><br><span data-ttu-id="57f70-890">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-890">
         - DocumentEvents</span></span><br><span data-ttu-id="57f70-891">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="57f70-891">
         - File</span></span><br><span data-ttu-id="57f70-892">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="57f70-892">
         - PdfFile</span></span><br><span data-ttu-id="57f70-893">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="57f70-893">
         - Selection</span></span><br><span data-ttu-id="57f70-894">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="57f70-894">
         - Settings</span></span><br><span data-ttu-id="57f70-895">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-895">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="57f70-896">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="57f70-896">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="57f70-897">OneNote</span><span class="sxs-lookup"><span data-stu-id="57f70-897">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="57f70-898">Plataforma</span><span class="sxs-lookup"><span data-stu-id="57f70-898">Platform</span></span></th>
    <th><span data-ttu-id="57f70-899">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="57f70-899">Extension points</span></span></th>
    <th><span data-ttu-id="57f70-900">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="57f70-900">API requirement sets</span></span></th>
    <th><span data-ttu-id="57f70-901"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="57f70-901"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-902">Office na Web</span><span class="sxs-lookup"><span data-stu-id="57f70-902">Office on the web</span></span></td>
    <td> <span data-ttu-id="57f70-903">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="57f70-903">- Content</span></span><br><span data-ttu-id="57f70-904">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="57f70-904">
         - TaskPane</span></span><br><span data-ttu-id="57f70-905">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="57f70-905">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="57f70-906">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-906">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="57f70-907">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-907">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="57f70-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="57f70-909">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="57f70-909">- DocumentEvents</span></span><br><span data-ttu-id="57f70-910">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-910">
         - HtmlCoercion</span></span><br><span data-ttu-id="57f70-911">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="57f70-911">
         - Settings</span></span><br><span data-ttu-id="57f70-912">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-912">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="57f70-913">Project</span><span class="sxs-lookup"><span data-stu-id="57f70-913">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="57f70-914">Plataforma</span><span class="sxs-lookup"><span data-stu-id="57f70-914">Platform</span></span></th>
    <th><span data-ttu-id="57f70-915">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="57f70-915">Extension points</span></span></th>
    <th><span data-ttu-id="57f70-916">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="57f70-916">API requirement sets</span></span></th>
    <th><span data-ttu-id="57f70-917"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="57f70-917"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-918">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="57f70-918">Office 2019 on Windows</span></span><br><span data-ttu-id="57f70-919">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="57f70-919">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="57f70-920">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="57f70-920">- TaskPane</span></span></td>
    <td> <span data-ttu-id="57f70-921">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-921">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="57f70-922">- Seleção</span><span class="sxs-lookup"><span data-stu-id="57f70-922">- Selection</span></span><br><span data-ttu-id="57f70-923">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-923">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-924">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="57f70-924">Office 2016 on Windows</span></span><br><span data-ttu-id="57f70-925">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="57f70-925">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="57f70-926">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="57f70-926">- TaskPane</span></span></td>
    <td> <span data-ttu-id="57f70-927">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-927">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="57f70-928">- Seleção</span><span class="sxs-lookup"><span data-stu-id="57f70-928">- Selection</span></span><br><span data-ttu-id="57f70-929">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-929">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="57f70-930">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="57f70-930">Office 2013 on Windows</span></span><br><span data-ttu-id="57f70-931">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="57f70-931">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="57f70-932">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="57f70-932">- TaskPane</span></span></td>
    <td> <span data-ttu-id="57f70-933">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="57f70-933">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="57f70-934">- Seleção</span><span class="sxs-lookup"><span data-stu-id="57f70-934">- Selection</span></span><br><span data-ttu-id="57f70-935">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="57f70-935">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="57f70-936">Confira também</span><span class="sxs-lookup"><span data-stu-id="57f70-936">See also</span></span>

- [<span data-ttu-id="57f70-937">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="57f70-937">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="57f70-938">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="57f70-938">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="57f70-939">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="57f70-939">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="57f70-940">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="57f70-940">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="57f70-941">Documentação de Referência da API</span><span class="sxs-lookup"><span data-stu-id="57f70-941">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="57f70-942">Histórico de atualizações do Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="57f70-942">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="57f70-943">Histórico de atualizações do Office 2016 e 2019 (Clique para Executar)</span><span class="sxs-lookup"><span data-stu-id="57f70-943">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="57f70-944">Histórico de atualizações do Office 2013 (clique para executar)</span><span class="sxs-lookup"><span data-stu-id="57f70-944">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="57f70-945">Histórico de atualizações do Office 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="57f70-945">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="57f70-946">Histórico de atualizações do Outlook 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="57f70-946">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="57f70-947">Histórico de atualizações do Office para Mac</span><span class="sxs-lookup"><span data-stu-id="57f70-947">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="57f70-948">Criando Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="57f70-948">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)