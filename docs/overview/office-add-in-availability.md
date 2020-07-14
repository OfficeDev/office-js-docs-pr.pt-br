---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos com suporte para o Excel, OneNote, Outlook, PowerPoint, Project e Word.
ms.date: 07/07/2020
localization_priority: Priority
ms.openlocfilehash: 3e6124373295875d607d86e0c311a34982b2c046
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094110"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="6b80d-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="6b80d-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="6b80d-104">Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="6b80d-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="6b80d-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que são compatíveis atualmente com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="6b80d-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="6b80d-106">A versão inicial do Office 2016 instalada por meio do MSI apenas contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="6b80d-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="6b80d-107">Para saber mais sobre o histórico de atualizações de várias versões do Office, confira a seção[Confira também](#see-also).</span><span class="sxs-lookup"><span data-stu-id="6b80d-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="6b80d-108">Excel</span><span class="sxs-lookup"><span data-stu-id="6b80d-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="6b80d-109">Plataforma</span><span class="sxs-lookup"><span data-stu-id="6b80d-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="6b80d-110">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="6b80d-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="6b80d-111">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="6b80d-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="6b80d-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="6b80d-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-113">Office na Web</span><span class="sxs-lookup"><span data-stu-id="6b80d-113">Office on the web</span></span></td>
    <td><span data-ttu-id="6b80d-114">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="6b80d-114">
      - TaskPane</span></span><br><span data-ttu-id="6b80d-115">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6b80d-115">
      - Content</span></span><br><span data-ttu-id="6b80d-116">
      - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="6b80d-116">
      - Custom Functions</span></span><br><span data-ttu-id="6b80d-117">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-117">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-118">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-118">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6b80d-119">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-119">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="6b80d-120">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-120">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="6b80d-121">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-121">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="6b80d-122">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-122">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="6b80d-123">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-123">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="6b80d-124">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-124">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="6b80d-125">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-125">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="6b80d-126">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-126">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="6b80d-127">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-127">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="6b80d-128">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-128">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="6b80d-129">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-129">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="6b80d-130">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-130">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-131">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-131">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="6b80d-132">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-132">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="6b80d-133">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-133">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="6b80d-134">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-134">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="6b80d-135">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-135">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="6b80d-136">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-136">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="6b80d-137">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-137">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="6b80d-138">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-138">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="6b80d-139">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-139">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="6b80d-140">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-140">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="6b80d-141">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-141">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="6b80d-142">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-142">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-143">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="6b80d-143">Office on Windows</span></span><br><span data-ttu-id="6b80d-144">(Conectado à assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="6b80d-144">(connected to Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="6b80d-145">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="6b80d-145">
      - TaskPane</span></span><br><span data-ttu-id="6b80d-146">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6b80d-146">
      - Content</span></span><br><span data-ttu-id="6b80d-147">
      - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="6b80d-147">
      - Custom Functions</span></span><br><span data-ttu-id="6b80d-148">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-148">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-149">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-149">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6b80d-150">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-150">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="6b80d-151">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-151">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="6b80d-152">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-152">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="6b80d-153">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-153">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="6b80d-154">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-154">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="6b80d-155">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-155">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="6b80d-156">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-156">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="6b80d-157">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-157">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="6b80d-158">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-158">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="6b80d-159">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-159">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="6b80d-160">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-160">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6b80d-161">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-161">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="6b80d-162">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-162">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-163">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-163">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="6b80d-164">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-164">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="6b80d-165">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-165">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="6b80d-166">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-166">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="6b80d-167">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-167">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="6b80d-168">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-168">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="6b80d-169">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-169">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="6b80d-170">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-170">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="6b80d-171">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-171">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="6b80d-172">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-172">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="6b80d-173">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-173">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="6b80d-174">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-174">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-175">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="6b80d-175">Office 2019 on Windows</span></span><br><span data-ttu-id="6b80d-176">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="6b80d-176">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="6b80d-177">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="6b80d-177">
      - TaskPane</span></span><br><span data-ttu-id="6b80d-178">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6b80d-178">
      - Content</span></span><br><span data-ttu-id="6b80d-179">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-179">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-180">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-180">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6b80d-181">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-181">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="6b80d-182">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-182">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="6b80d-183">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-183">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="6b80d-184">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-184">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="6b80d-185">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-185">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="6b80d-186">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-186">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="6b80d-187">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-187">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="6b80d-188">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-188">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6b80d-189">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-189">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-190">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-190">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="6b80d-191">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-191">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="6b80d-192">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-192">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="6b80d-193">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-193">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="6b80d-194">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-194">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="6b80d-195">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-195">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="6b80d-196">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-196">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="6b80d-197">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-197">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="6b80d-198">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-198">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="6b80d-199">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-199">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="6b80d-200">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-200">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="6b80d-201">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-201">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-202">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="6b80d-202">Office 2016 on Windows</span></span><br><span data-ttu-id="6b80d-203">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="6b80d-203">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="6b80d-204">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="6b80d-204">
      - TaskPane</span></span><br><span data-ttu-id="6b80d-205">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6b80d-205">
      - Content</span></span> </td>
    <td><span data-ttu-id="6b80d-206">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-206">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6b80d-207">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="6b80d-207">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="6b80d-208">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-208">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-209">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-209">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="6b80d-210">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-210">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="6b80d-211">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-211">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="6b80d-212">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-212">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="6b80d-213">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-213">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="6b80d-214">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-214">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="6b80d-215">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-215">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="6b80d-216">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-216">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="6b80d-217">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-217">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="6b80d-218">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-218">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="6b80d-219">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-219">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="6b80d-220">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-220">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-221">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="6b80d-221">Office 2013 on Windows</span></span><br><span data-ttu-id="6b80d-222">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="6b80d-222">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="6b80d-223">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="6b80d-223">
      - TaskPane</span></span><br><span data-ttu-id="6b80d-224">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6b80d-224">
      - Content</span></span> </td>
    <td><span data-ttu-id="6b80d-225">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="6b80d-225">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="6b80d-226">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-226">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-227">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-227">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="6b80d-228">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-228">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="6b80d-229">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-229">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="6b80d-230">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-230">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="6b80d-231">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-231">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="6b80d-232">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-232">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="6b80d-233">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-233">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="6b80d-234">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-234">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="6b80d-235">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-235">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="6b80d-236">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-236">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="6b80d-237">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-237">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-238">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="6b80d-238">Office on iPad</span></span><br><span data-ttu-id="6b80d-239">(Conectado à assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="6b80d-239">(connected to Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="6b80d-240">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="6b80d-240">
      - TaskPane</span></span><br><span data-ttu-id="6b80d-241">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6b80d-241">
      - Content</span></span> </td>
    <td><span data-ttu-id="6b80d-242">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-242">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6b80d-243">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-243">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="6b80d-244">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-244">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="6b80d-245">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-245">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="6b80d-246">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-246">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="6b80d-247">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-247">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="6b80d-248">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-248">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="6b80d-249">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-249">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="6b80d-250">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-250">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="6b80d-251">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-251">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="6b80d-252">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-252">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="6b80d-253">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-253">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6b80d-254">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-254">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-255">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-255">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="6b80d-256">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-256">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="6b80d-257">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-257">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="6b80d-258">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-258">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="6b80d-259">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-259">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="6b80d-260">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-260">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="6b80d-261">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-261">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="6b80d-262">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-262">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="6b80d-263">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-263">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="6b80d-264">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-264">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="6b80d-265">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-265">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-266">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="6b80d-266">Office on Mac</span></span><br><span data-ttu-id="6b80d-267">(Conectado à assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="6b80d-267">(connected to Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="6b80d-268">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="6b80d-268">
      - TaskPane</span></span><br><span data-ttu-id="6b80d-269">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6b80d-269">
      - Content</span></span><br><span data-ttu-id="6b80d-270">
      - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="6b80d-270">
      - Custom Functions</span></span><br><span data-ttu-id="6b80d-271">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-271">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-272">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-272">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6b80d-273">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-273">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="6b80d-274">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-274">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="6b80d-275">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-275">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="6b80d-276">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-276">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="6b80d-277">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-277">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="6b80d-278">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-278">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="6b80d-279">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-279">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="6b80d-280">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-280">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="6b80d-281">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-281">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="6b80d-282">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-282">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="6b80d-283">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-283">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6b80d-284">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-284">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="6b80d-285">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-285">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-286">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-286">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="6b80d-287">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-287">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="6b80d-288">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-288">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="6b80d-289">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-289">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="6b80d-290">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-290">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="6b80d-291">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-291">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="6b80d-292">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-292">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="6b80d-293">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-293">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="6b80d-294">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-294">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="6b80d-295">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-295">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="6b80d-296">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-296">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="6b80d-297">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-297">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="6b80d-298">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-298">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-299">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="6b80d-299">Office 2019 on Mac</span></span><br><span data-ttu-id="6b80d-300">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="6b80d-300">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="6b80d-301">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="6b80d-301">
      - TaskPane</span></span><br><span data-ttu-id="6b80d-302">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6b80d-302">
      - Content</span></span><br><span data-ttu-id="6b80d-303">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-303">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-304">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-304">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6b80d-305">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-305">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="6b80d-306">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-306">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="6b80d-307">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-307">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="6b80d-308">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-308">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="6b80d-309">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-309">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="6b80d-310">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-310">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="6b80d-311">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-311">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="6b80d-312">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-312">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6b80d-313">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-313">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-314">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-314">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="6b80d-315">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-315">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="6b80d-316">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-316">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="6b80d-317">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-317">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="6b80d-318">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-318">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="6b80d-319">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-319">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="6b80d-320">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-320">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="6b80d-321">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-321">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="6b80d-322">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-322">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="6b80d-323">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-323">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="6b80d-324">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-324">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="6b80d-325">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-325">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="6b80d-326">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-326">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-327">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="6b80d-327">Office 2016 on Mac</span></span><br><span data-ttu-id="6b80d-328">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="6b80d-328">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="6b80d-329">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="6b80d-329">
      - TaskPane</span></span><br><span data-ttu-id="6b80d-330">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6b80d-330">
      - Content</span></span> </td>
    <td><span data-ttu-id="6b80d-331">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-331">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6b80d-332">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="6b80d-332">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="6b80d-333">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-333">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-334">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-334">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="6b80d-335">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-335">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="6b80d-336">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-336">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="6b80d-337">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-337">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="6b80d-338">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-338">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="6b80d-339">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-339">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="6b80d-340">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-340">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="6b80d-341">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-341">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="6b80d-342">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-342">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="6b80d-343">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-343">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="6b80d-344">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-344">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="6b80d-345">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-345">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="6b80d-346">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-346">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
</table>

<span data-ttu-id="6b80d-347">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="6b80d-347">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="6b80d-348">Funções personalizadas (somente Excel)</span><span class="sxs-lookup"><span data-stu-id="6b80d-348">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="6b80d-349">Plataforma</span><span class="sxs-lookup"><span data-stu-id="6b80d-349">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="6b80d-350">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="6b80d-350">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="6b80d-351">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="6b80d-351">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="6b80d-352"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="6b80d-352"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-353">Office na Web</span><span class="sxs-lookup"><span data-stu-id="6b80d-353">Office on the web</span></span></td>
    <td><span data-ttu-id="6b80d-354">- Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="6b80d-354">- Custom Functions</span></span></td>
    <td><span data-ttu-id="6b80d-355">- <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-355">- <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-356">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="6b80d-356">Office on Windows</span></span><br><span data-ttu-id="6b80d-357">(Conectado à assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="6b80d-357">(connected to Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="6b80d-358">- Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="6b80d-358">- Custom Functions</span></span></td>
    <td><span data-ttu-id="6b80d-359">- <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-359">- <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-360">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="6b80d-360">Office on Mac</span></span><br><span data-ttu-id="6b80d-361">(Conectado à assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="6b80d-361">(connected to Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="6b80d-362">- Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="6b80d-362">- Custom Functions</span></span></td>
    <td><span data-ttu-id="6b80d-363">- <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-363">- <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td></td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="6b80d-364">Outlook</span><span class="sxs-lookup"><span data-stu-id="6b80d-364">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="6b80d-365">Plataforma</span><span class="sxs-lookup"><span data-stu-id="6b80d-365">Platform</span></span></th>
    <th><span data-ttu-id="6b80d-366">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="6b80d-366">Extension points</span></span></th>
    <th><span data-ttu-id="6b80d-367">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="6b80d-367">API requirement sets</span></span></th>
    <th><span data-ttu-id="6b80d-368"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="6b80d-368"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-369">Office na Web</span><span class="sxs-lookup"><span data-stu-id="6b80d-369">Office on the web</span></span><br><span data-ttu-id="6b80d-370">(moderno)</span><span class="sxs-lookup"><span data-stu-id="6b80d-370">(modern)</span></span></td>
    <td><span data-ttu-id="6b80d-371">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-371">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="6b80d-372">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-372">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="6b80d-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="6b80d-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="6b80d-375">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-375">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6b80d-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6b80d-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6b80d-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6b80d-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="6b80d-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="6b80d-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="6b80d-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-384">Não disponível</span><span class="sxs-lookup"><span data-stu-id="6b80d-384">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-385">Office na Web</span><span class="sxs-lookup"><span data-stu-id="6b80d-385">Office on the web</span></span><br><span data-ttu-id="6b80d-386">(clássico)</span><span class="sxs-lookup"><span data-stu-id="6b80d-386">(classic)</span></span></td>
    <td><span data-ttu-id="6b80d-387">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-387">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="6b80d-388">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-388">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="6b80d-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="6b80d-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="6b80d-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6b80d-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6b80d-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6b80d-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6b80d-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="6b80d-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-398">Não disponível</span><span class="sxs-lookup"><span data-stu-id="6b80d-398">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-399">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="6b80d-399">Office on Windows</span></span><br><span data-ttu-id="6b80d-400">(Conectado à assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="6b80d-400">(connected to Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="6b80d-401">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-401">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="6b80d-402">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-402">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="6b80d-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="6b80d-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="6b80d-405">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-405">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="6b80d-406">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Módulos</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-406">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6b80d-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6b80d-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6b80d-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6b80d-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="6b80d-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="6b80d-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="6b80d-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-415">Não disponível</span><span class="sxs-lookup"><span data-stu-id="6b80d-415">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-416">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="6b80d-416">Office 2019 on Windows</span></span><br><span data-ttu-id="6b80d-417">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="6b80d-417">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="6b80d-418">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-418">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="6b80d-419">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-419">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="6b80d-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="6b80d-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="6b80d-422">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-422">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="6b80d-423">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Módulos</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-423">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6b80d-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6b80d-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6b80d-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6b80d-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="6b80d-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="6b80d-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-431">Não disponível</span><span class="sxs-lookup"><span data-stu-id="6b80d-431">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-432">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="6b80d-432">Office 2016 on Windows</span></span><br><span data-ttu-id="6b80d-433">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="6b80d-433">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="6b80d-434">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-434">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="6b80d-435">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-435">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="6b80d-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="6b80d-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="6b80d-438">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-438">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="6b80d-439">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Módulos</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-439">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6b80d-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6b80d-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6b80d-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*
    </span></span></td>
    <td><span data-ttu-id="6b80d-444">Não disponível</span><span class="sxs-lookup"><span data-stu-id="6b80d-444">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-445">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="6b80d-445">Office 2013 on Windows</span></span><br><span data-ttu-id="6b80d-446">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="6b80d-446">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="6b80d-447">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-447">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="6b80d-448">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-448">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="6b80d-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="6b80d-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br>
    </td>
    <td><span data-ttu-id="6b80d-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6b80d-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6b80d-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="6b80d-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="6b80d-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*
    </span></span></td>
    <td><span data-ttu-id="6b80d-455">Não disponível</span><span class="sxs-lookup"><span data-stu-id="6b80d-455">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-456">Office no iOS</span><span class="sxs-lookup"><span data-stu-id="6b80d-456">Office on iOS</span></span><br><span data-ttu-id="6b80d-457">(Conectado à assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="6b80d-457">(connected to Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="6b80d-458">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-458">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="6b80d-459">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-459">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6b80d-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6b80d-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6b80d-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6b80d-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-465">Não disponível</span><span class="sxs-lookup"><span data-stu-id="6b80d-465">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-466">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="6b80d-466">Office on Mac</span></span><br><span data-ttu-id="6b80d-467">(Conectado à assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="6b80d-467">(connected to Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="6b80d-468">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-468">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="6b80d-469">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-469">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="6b80d-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="6b80d-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="6b80d-472">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-472">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6b80d-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6b80d-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6b80d-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6b80d-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="6b80d-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="6b80d-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="6b80d-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-481">Não disponível</span><span class="sxs-lookup"><span data-stu-id="6b80d-481">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-482">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="6b80d-482">Office 2019 on Mac</span></span><br><span data-ttu-id="6b80d-483">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="6b80d-483">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="6b80d-484">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-484">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="6b80d-485">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-485">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="6b80d-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="6b80d-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="6b80d-488">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-488">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6b80d-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6b80d-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6b80d-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6b80d-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="6b80d-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-495">Não disponível</span><span class="sxs-lookup"><span data-stu-id="6b80d-495">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-496">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="6b80d-496">Office 2016 on Mac</span></span><br><span data-ttu-id="6b80d-497">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="6b80d-497">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="6b80d-498">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-498">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="6b80d-499">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-499">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="6b80d-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="6b80d-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="6b80d-502">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-502">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-503">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-503">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6b80d-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6b80d-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6b80d-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6b80d-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="6b80d-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-509">Não disponível</span><span class="sxs-lookup"><span data-stu-id="6b80d-509">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-510">Outlook no Android</span><span class="sxs-lookup"><span data-stu-id="6b80d-510">Office on Android</span></span><br><span data-ttu-id="6b80d-511">(Conectado à assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="6b80d-511">(connected to Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="6b80d-512">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-512">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="6b80d-513">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Organizador de compromissos (Redigir): reunião on-line (visualização)</span><span class="sxs-lookup"><span data-stu-id="6b80d-513">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Appointment Organizer (Compose): online meeting</a> (preview)</span></span><br><span data-ttu-id="6b80d-514">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-514">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-515">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-515">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6b80d-516">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-516">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6b80d-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6b80d-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6b80d-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-520">Não disponível</span><span class="sxs-lookup"><span data-stu-id="6b80d-520">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="6b80d-521">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="6b80d-521">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="6b80d-522">O suporte ao cliente para um conjunto de requisitos pode ser restringido pelo suporte do servidor Exchange.</span><span class="sxs-lookup"><span data-stu-id="6b80d-522">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="6b80d-523">Consulte [Conjuntos de requisitos da API JavaScript do Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) para obter detalhes sobre o intervalo de conjuntos de requisitos suportado pelo servidor Exchange e pelos clientes Outlook.</span><span class="sxs-lookup"><span data-stu-id="6b80d-523">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="6b80d-524">Word</span><span class="sxs-lookup"><span data-stu-id="6b80d-524">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="6b80d-525">Plataforma</span><span class="sxs-lookup"><span data-stu-id="6b80d-525">Platform</span></span></th>
    <th><span data-ttu-id="6b80d-526">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="6b80d-526">Extension points</span></span></th>
    <th><span data-ttu-id="6b80d-527">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="6b80d-527">API requirement sets</span></span></th>
    <th><span data-ttu-id="6b80d-528"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="6b80d-528"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-529">Office na Web</span><span class="sxs-lookup"><span data-stu-id="6b80d-529">Office on the web</span></span></td>
    <td><span data-ttu-id="6b80d-530">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="6b80d-530">
      - TaskPane</span></span><br><span data-ttu-id="6b80d-531">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-531">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-532">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-532">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="6b80d-533">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-533">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="6b80d-534">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-534">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="6b80d-535">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-535">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6b80d-536">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-536">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="6b80d-537">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-537">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-538">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-538">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="6b80d-539">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-539">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="6b80d-540">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-540">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="6b80d-541">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-541">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="6b80d-542">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-542">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="6b80d-543">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-543">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="6b80d-544">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-544">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="6b80d-545">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-545">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="6b80d-546">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-546">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="6b80d-547">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-547">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="6b80d-548">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-548">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="6b80d-549">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-549">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="6b80d-550">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-550">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="6b80d-551">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-551">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="6b80d-552">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-552">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="6b80d-553">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-553">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-554">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="6b80d-554">Office on Windows</span></span><br><span data-ttu-id="6b80d-555">(Conectado à assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="6b80d-555">(connected to Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="6b80d-556">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="6b80d-556">
      - TaskPane</span></span><br><span data-ttu-id="6b80d-557">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-557">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-558">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-558">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="6b80d-559">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-559">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="6b80d-560">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-560">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="6b80d-561">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-561">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6b80d-562">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-562">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="6b80d-563">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-563">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-564">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-564">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="6b80d-565">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-565">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="6b80d-566">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-566">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="6b80d-567">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-567">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="6b80d-568">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-568">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="6b80d-569">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-569">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="6b80d-570">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-570">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="6b80d-571">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-571">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="6b80d-572">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-572">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="6b80d-573">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-573">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="6b80d-574">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-574">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="6b80d-575">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-575">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="6b80d-576">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-576">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="6b80d-577">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-577">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="6b80d-578">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-578">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="6b80d-579">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-579">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="6b80d-580">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-580">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-581">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="6b80d-581">Office 2019 on Windows</span></span><br><span data-ttu-id="6b80d-582">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="6b80d-582">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="6b80d-583">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="6b80d-583">
      - TaskPane</span></span><br><span data-ttu-id="6b80d-584">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-584">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-585">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-585">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="6b80d-586">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-586">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="6b80d-587">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-587">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="6b80d-588">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-588">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6b80d-589">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-589">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-590">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-590">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="6b80d-591">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-591">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="6b80d-592">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-592">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="6b80d-593">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-593">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="6b80d-594">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-594">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="6b80d-595">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-595">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="6b80d-596">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-596">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="6b80d-597">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-597">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="6b80d-598">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-598">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="6b80d-599">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-599">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="6b80d-600">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-600">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="6b80d-601">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-601">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="6b80d-602">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-602">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="6b80d-603">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-603">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="6b80d-604">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-604">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="6b80d-605">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-605">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="6b80d-606">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-606">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-607">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="6b80d-607">Office 2016 on Windows</span></span><br><span data-ttu-id="6b80d-608">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="6b80d-608">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="6b80d-609">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="6b80d-609">- TaskPane</span></span></td>
    <td><span data-ttu-id="6b80d-610">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-610">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="6b80d-611">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="6b80d-611">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="6b80d-612">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-612">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-613">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-613">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="6b80d-614">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-614">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="6b80d-615">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-615">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="6b80d-616">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-616">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="6b80d-617">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-617">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="6b80d-618">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-618">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="6b80d-619">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-619">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="6b80d-620">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-620">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="6b80d-621">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-621">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="6b80d-622">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-622">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="6b80d-623">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-623">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="6b80d-624">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-624">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="6b80d-625">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-625">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="6b80d-626">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-626">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="6b80d-627">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-627">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="6b80d-628">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-628">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="6b80d-629">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-629">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-630">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="6b80d-630">Office 2013 on Windows</span></span><br><span data-ttu-id="6b80d-631">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="6b80d-631">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="6b80d-632">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="6b80d-632">- TaskPane</span></span></td>
    <td><span data-ttu-id="6b80d-633">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="6b80d-633">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="6b80d-634">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-634">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-635">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-635">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="6b80d-636">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-636">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="6b80d-637">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-637">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="6b80d-638">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-638">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="6b80d-639">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-639">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="6b80d-640">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-640">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="6b80d-641">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-641">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="6b80d-642">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-642">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="6b80d-643">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-643">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="6b80d-644">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-644">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="6b80d-645">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-645">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="6b80d-646">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-646">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="6b80d-647">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-647">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="6b80d-648">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-648">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="6b80d-649">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-649">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="6b80d-650">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-650">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="6b80d-651">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-651">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-652">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="6b80d-652">Office on iPad</span></span><br><span data-ttu-id="6b80d-653">(Conectado à assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="6b80d-653">(connected to Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="6b80d-654">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="6b80d-654">- TaskPane</span></span></td>
    <td><span data-ttu-id="6b80d-655">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-655">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="6b80d-656">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-656">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="6b80d-657">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-657">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="6b80d-658">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-658">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6b80d-659">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-659">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-660">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-660">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="6b80d-661">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-661">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="6b80d-662">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-662">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="6b80d-663">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-663">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="6b80d-664">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-664">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="6b80d-665">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-665">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="6b80d-666">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-666">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="6b80d-667">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-667">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="6b80d-668">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-668">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="6b80d-669">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-669">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="6b80d-670">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-670">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="6b80d-671">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-671">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="6b80d-672">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-672">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="6b80d-673">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-673">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="6b80d-674">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-674">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="6b80d-675">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-675">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="6b80d-676">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-676">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-677">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="6b80d-677">Office on Mac</span></span><br><span data-ttu-id="6b80d-678">(Conectado à assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="6b80d-678">(connected to Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="6b80d-679">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="6b80d-679">
      - TaskPane</span></span><br><span data-ttu-id="6b80d-680">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-680">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-681">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-681">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="6b80d-682">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-682">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="6b80d-683">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-683">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="6b80d-684">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-684">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6b80d-685">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-685">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="6b80d-686">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-686">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-687">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-687">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="6b80d-688">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-688">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="6b80d-689">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-689">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="6b80d-690">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-690">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="6b80d-691">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-691">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="6b80d-692">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-692">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="6b80d-693">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-693">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="6b80d-694">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-694">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="6b80d-695">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-695">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="6b80d-696">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-696">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="6b80d-697">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-697">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="6b80d-698">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-698">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="6b80d-699">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-699">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="6b80d-700">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-700">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="6b80d-701">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-701">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="6b80d-702">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-702">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="6b80d-703">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-703">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-704">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="6b80d-704">Office 2019 on Mac</span></span><br><span data-ttu-id="6b80d-705">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="6b80d-705">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="6b80d-706">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="6b80d-706">
      - TaskPane</span></span><br><span data-ttu-id="6b80d-707">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-707">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-708">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-708">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="6b80d-709">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-709">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="6b80d-710">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-710">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="6b80d-711">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-711">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6b80d-712">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-712">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-713">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-713">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="6b80d-714">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-714">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="6b80d-715">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-715">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="6b80d-716">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-716">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="6b80d-717">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-717">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="6b80d-718">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-718">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="6b80d-719">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-719">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="6b80d-720">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-720">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="6b80d-721">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-721">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="6b80d-722">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-722">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="6b80d-723">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-723">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="6b80d-724">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-724">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="6b80d-725">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-725">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="6b80d-726">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-726">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="6b80d-727">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-727">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="6b80d-728">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-728">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="6b80d-729">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-729">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-730">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="6b80d-730">Office 2016 on Mac</span></span><br><span data-ttu-id="6b80d-731">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="6b80d-731">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="6b80d-732">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="6b80d-732">- TaskPane</span></span></td>
    <td><span data-ttu-id="6b80d-733">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-733">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="6b80d-734">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="6b80d-734">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="6b80d-735">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-735">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-736">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-736">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="6b80d-737">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-737">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="6b80d-738">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-738">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="6b80d-739">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-739">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="6b80d-740">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-740">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="6b80d-741">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-741">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="6b80d-742">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-742">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="6b80d-743">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-743">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="6b80d-744">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-744">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="6b80d-745">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-745">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="6b80d-746">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-746">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="6b80d-747">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-747">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="6b80d-748">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-748">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="6b80d-749">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-749">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="6b80d-750">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-750">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="6b80d-751">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-751">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="6b80d-752">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-752">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
</table>

<span data-ttu-id="6b80d-753">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="6b80d-753">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="6b80d-754">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="6b80d-754">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="6b80d-755">Plataforma</span><span class="sxs-lookup"><span data-stu-id="6b80d-755">Platform</span></span></th>
    <th><span data-ttu-id="6b80d-756">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="6b80d-756">Extension points</span></span></th>
    <th><span data-ttu-id="6b80d-757">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="6b80d-757">API requirement sets</span></span></th>
    <th><span data-ttu-id="6b80d-758"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="6b80d-758"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-759">Office na Web</span><span class="sxs-lookup"><span data-stu-id="6b80d-759">Office on the web</span></span></td>
    <td><span data-ttu-id="6b80d-760">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6b80d-760">
      - Content</span></span><br><span data-ttu-id="6b80d-761">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="6b80d-761">
      - TaskPane</span></span><br><span data-ttu-id="6b80d-762">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-762">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-763">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-763">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="6b80d-764">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-764">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6b80d-765">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-765">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="6b80d-766">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-766">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-767">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-767">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="6b80d-768">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-768">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="6b80d-769">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-769">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="6b80d-770">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-770">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="6b80d-771">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-771">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="6b80d-772">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-772">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="6b80d-773">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-773">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="6b80d-774">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-774">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-775">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="6b80d-775">Office on Windows</span></span><br><span data-ttu-id="6b80d-776">(Conectado à assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="6b80d-776">(connected to Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="6b80d-777">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6b80d-777">
      - Content</span></span><br><span data-ttu-id="6b80d-778">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="6b80d-778">
      - TaskPane</span></span><br><span data-ttu-id="6b80d-779">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-779">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-780">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-780">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="6b80d-781">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-781">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6b80d-782">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-782">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="6b80d-783">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-783">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-784">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-784">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="6b80d-785">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-785">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="6b80d-786">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-786">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="6b80d-787">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-787">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="6b80d-788">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-788">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="6b80d-789">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-789">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="6b80d-790">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-790">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="6b80d-791">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-791">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-792">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="6b80d-792">Office 2019 on Windows</span></span><br><span data-ttu-id="6b80d-793">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="6b80d-793">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="6b80d-794">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6b80d-794">
      - Content</span></span><br><span data-ttu-id="6b80d-795">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="6b80d-795">
      - TaskPane</span></span><br><span data-ttu-id="6b80d-796">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-796">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-797">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-797">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6b80d-798">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-798">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-799">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-799">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="6b80d-800">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-800">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="6b80d-801">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-801">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="6b80d-802">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-802">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="6b80d-803">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-803">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="6b80d-804">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-804">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="6b80d-805">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-805">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="6b80d-806">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-806">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-807">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="6b80d-807">Office 2016 on Windows</span></span><br><span data-ttu-id="6b80d-808">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="6b80d-808">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="6b80d-809">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6b80d-809">
      - Content</span></span><br><span data-ttu-id="6b80d-810">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="6b80d-810">
      - TaskPane</span></span> </td>
    <td><span data-ttu-id="6b80d-811">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="6b80d-811">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="6b80d-812">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-812">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-813">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-813">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="6b80d-814">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-814">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="6b80d-815">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-815">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="6b80d-816">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-816">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="6b80d-817">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-817">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="6b80d-818">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-818">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="6b80d-819">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-819">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="6b80d-820">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-820">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-821">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="6b80d-821">Office 2013 on Windows</span></span><br><span data-ttu-id="6b80d-822">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="6b80d-822">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="6b80d-823">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6b80d-823">
      - Content</span></span><br><span data-ttu-id="6b80d-824">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="6b80d-824">
      - TaskPane</span></span> </td>
    <td><span data-ttu-id="6b80d-825">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="6b80d-825">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="6b80d-826">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-826">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-827">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-827">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="6b80d-828">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-828">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="6b80d-829">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-829">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="6b80d-830">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-830">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="6b80d-831">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-831">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="6b80d-832">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-832">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="6b80d-833">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-833">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="6b80d-834">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-834">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-835">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="6b80d-835">Office on iPad</span></span><br><span data-ttu-id="6b80d-836">(Conectado à assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="6b80d-836">(connected to Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="6b80d-837">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6b80d-837">
      - Content</span></span><br><span data-ttu-id="6b80d-838">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="6b80d-838">
      - TaskPane</span></span> </td>
    <td><span data-ttu-id="6b80d-839">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-839">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="6b80d-840">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-840">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6b80d-841">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-841">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-842">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-842">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="6b80d-843">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-843">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="6b80d-844">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-844">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="6b80d-845">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-845">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="6b80d-846">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-846">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="6b80d-847">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-847">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="6b80d-848">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-848">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="6b80d-849">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-849">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-850">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="6b80d-850">Office on Mac</span></span><br><span data-ttu-id="6b80d-851">(Conectado à assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="6b80d-851">(connected to Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="6b80d-852">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6b80d-852">
      - Content</span></span><br><span data-ttu-id="6b80d-853">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="6b80d-853">
      - TaskPane</span></span><br><span data-ttu-id="6b80d-854">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-854">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-855">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-855">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="6b80d-856">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-856">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6b80d-857">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-857">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="6b80d-858">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-858">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-859">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-859">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="6b80d-860">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-860">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="6b80d-861">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-861">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="6b80d-862">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-862">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="6b80d-863">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-863">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="6b80d-864">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-864">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="6b80d-865">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-865">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="6b80d-866">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-866">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-867">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="6b80d-867">Office 2019 on Mac</span></span><br><span data-ttu-id="6b80d-868">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="6b80d-868">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="6b80d-869">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6b80d-869">
      - Content</span></span><br><span data-ttu-id="6b80d-870">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="6b80d-870">
      - TaskPane</span></span><br><span data-ttu-id="6b80d-871">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-871">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-872">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-872">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6b80d-873">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-873">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-874">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-874">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="6b80d-875">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-875">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="6b80d-876">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-876">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="6b80d-877">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-877">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="6b80d-878">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-878">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="6b80d-879">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-879">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="6b80d-880">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-880">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="6b80d-881">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-881">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-882">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="6b80d-882">Office 2016 on Mac</span></span><br><span data-ttu-id="6b80d-883">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="6b80d-883">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="6b80d-884">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6b80d-884">
      - Content</span></span><br><span data-ttu-id="6b80d-885">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="6b80d-885">
      - TaskPane</span></span> </td>
    <td><span data-ttu-id="6b80d-886">
       - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="6b80d-886">
       - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="6b80d-887">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-887">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-888">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-888">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="6b80d-889">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-889">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="6b80d-890">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-890">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="6b80d-891">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-891">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="6b80d-892">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-892">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="6b80d-893">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-893">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="6b80d-894">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-894">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="6b80d-895">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-895">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
</table>

<span data-ttu-id="6b80d-896">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="6b80d-896">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="6b80d-897">OneNote</span><span class="sxs-lookup"><span data-stu-id="6b80d-897">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="6b80d-898">Plataforma</span><span class="sxs-lookup"><span data-stu-id="6b80d-898">Platform</span></span></th>
    <th><span data-ttu-id="6b80d-899">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="6b80d-899">Extension points</span></span></th>
    <th><span data-ttu-id="6b80d-900">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="6b80d-900">API requirement sets</span></span></th>
    <th><span data-ttu-id="6b80d-901"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="6b80d-901"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-902">Office na Web</span><span class="sxs-lookup"><span data-stu-id="6b80d-902">Office on the web</span></span></td>
    <td><span data-ttu-id="6b80d-903">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6b80d-903">
      - Content</span></span><br><span data-ttu-id="6b80d-904">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="6b80d-904">
      - TaskPane</span></span><br><span data-ttu-id="6b80d-905">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-905">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-906">
      - <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-906">
      - <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="6b80d-907">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-907">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6b80d-908">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-908">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="6b80d-909">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-909">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="6b80d-910">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-910">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="6b80d-911">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-911">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="6b80d-912">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-912">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="6b80d-913">Project</span><span class="sxs-lookup"><span data-stu-id="6b80d-913">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="6b80d-914">Plataforma</span><span class="sxs-lookup"><span data-stu-id="6b80d-914">Platform</span></span></th>
    <th><span data-ttu-id="6b80d-915">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="6b80d-915">Extension points</span></span></th>
    <th><span data-ttu-id="6b80d-916">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="6b80d-916">API requirement sets</span></span></th>
    <th><span data-ttu-id="6b80d-917"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="6b80d-917"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-918">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="6b80d-918">Office 2019 on Windows</span></span><br><span data-ttu-id="6b80d-919">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="6b80d-919">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="6b80d-920">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="6b80d-920">- TaskPane</span></span></td>
    <td><span data-ttu-id="6b80d-921">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-921">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="6b80d-922">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-922">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="6b80d-923">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-923">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-924">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="6b80d-924">Office 2016 on Windows</span></span><br><span data-ttu-id="6b80d-925">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="6b80d-925">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="6b80d-926">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="6b80d-926">- TaskPane</span></span></td>
    <td><span data-ttu-id="6b80d-927">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-927">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="6b80d-928">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-928">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="6b80d-929">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-929">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6b80d-930">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="6b80d-930">Office 2013 on Windows</span></span><br><span data-ttu-id="6b80d-931">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="6b80d-931">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="6b80d-932">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="6b80d-932">- TaskPane</span></span></td>
    <td><span data-ttu-id="6b80d-933">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-933">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="6b80d-934">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="6b80d-934">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="6b80d-935">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="6b80d-935">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="6b80d-936">Confira também</span><span class="sxs-lookup"><span data-stu-id="6b80d-936">See also</span></span>

- [<span data-ttu-id="6b80d-937">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="6b80d-937">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="6b80d-938">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="6b80d-938">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="6b80d-939">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="6b80d-939">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="6b80d-940">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="6b80d-940">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="6b80d-941">Documentação de Referência da API</span><span class="sxs-lookup"><span data-stu-id="6b80d-941">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="6b80d-942">Histórico de atualizações para Microsoft 365 Apps</span><span class="sxs-lookup"><span data-stu-id="6b80d-942">Update history for Microsoft 365 Apps</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="6b80d-943">Histórico de atualizações do Office 2016 e 2019 (Clique para Executar)</span><span class="sxs-lookup"><span data-stu-id="6b80d-943">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="6b80d-944">Histórico de atualizações do Office 2013 (clique para executar)</span><span class="sxs-lookup"><span data-stu-id="6b80d-944">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="6b80d-945">Histórico de atualizações do Office 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="6b80d-945">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="6b80d-946">Histórico de atualizações do Outlook 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="6b80d-946">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="6b80d-947">Histórico de atualizações do Office para Mac</span><span class="sxs-lookup"><span data-stu-id="6b80d-947">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="6b80d-948">Criando Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="6b80d-948">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)