---
title: Disponibilidade de aplicativos e plataformas do cliente Office para Suplementos do Office
description: Conjuntos de requisitos com suporte para o Excel, OneNote, Outlook, PowerPoint, Project e Word.
ms.date: 07/10/2020
localization_priority: Priority
ms.openlocfilehash: 4b78b320e36f14c774c03fff56ddef7d369fb13c
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292423"
---
# <a name="office-client-application-and-platform-availability-for-office-add-ins"></a><span data-ttu-id="3a005-103">Disponibilidade de aplicativos e plataformas do cliente Office para Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="3a005-103">Office client application and platform availability for Office Add-ins</span></span>

<span data-ttu-id="3a005-104">Seu Suplemento do Office pode depender de um aplicativo específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="3a005-104">To work as expected, your Office Add-in might depend on a specific Office application, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="3a005-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que são compatíveis atualmente com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="3a005-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="3a005-106">A versão inicial do Office 2016 instalada por meio do MSI apenas contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="3a005-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="3a005-107">Para saber mais sobre o histórico de atualizações de várias versões do Office, confira a seção[Confira também](#see-also).</span><span class="sxs-lookup"><span data-stu-id="3a005-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="3a005-108">Excel</span><span class="sxs-lookup"><span data-stu-id="3a005-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="3a005-109">Plataforma</span><span class="sxs-lookup"><span data-stu-id="3a005-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="3a005-110">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="3a005-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="3a005-111">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="3a005-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="3a005-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="3a005-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-113">Office na Web</span><span class="sxs-lookup"><span data-stu-id="3a005-113">Office on the web</span></span></td>
    <td><span data-ttu-id="3a005-114">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="3a005-114">
      - TaskPane</span></span><br><span data-ttu-id="3a005-115">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3a005-115">
      - Content</span></span><br><span data-ttu-id="3a005-116">
      - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="3a005-116">
      - Custom Functions</span></span><br><span data-ttu-id="3a005-117">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-117">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-118">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-118">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="3a005-119">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3a005-119">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="3a005-120">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3a005-120">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="3a005-121">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3a005-121">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="3a005-122">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3a005-122">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="3a005-123">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3a005-123">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="3a005-124">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3a005-124">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="3a005-125">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="3a005-125">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="3a005-126">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="3a005-126">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="3a005-127">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="3a005-127">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="3a005-128">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="3a005-128">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="3a005-129">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="3a005-129">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="3a005-130">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-130">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-131">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-131">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="3a005-132">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-132">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="3a005-133">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-133">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="3a005-134">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="3a005-134">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="3a005-135">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-135">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="3a005-136">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-136">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="3a005-137">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="3a005-137">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="3a005-138">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="3a005-138">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="3a005-139">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-139">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="3a005-140">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-140">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="3a005-141">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-141">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="3a005-142">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-142">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-143">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="3a005-143">Office on Windows</span></span><br><span data-ttu-id="3a005-144">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="3a005-144">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="3a005-145">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="3a005-145">
      - TaskPane</span></span><br><span data-ttu-id="3a005-146">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3a005-146">
      - Content</span></span><br><span data-ttu-id="3a005-147">
      - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="3a005-147">
      - Custom Functions</span></span><br><span data-ttu-id="3a005-148">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-148">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-149">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-149">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="3a005-150">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3a005-150">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="3a005-151">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3a005-151">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="3a005-152">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3a005-152">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="3a005-153">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3a005-153">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="3a005-154">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3a005-154">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="3a005-155">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3a005-155">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="3a005-156">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="3a005-156">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="3a005-157">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="3a005-157">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="3a005-158">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="3a005-158">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="3a005-159">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="3a005-159">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="3a005-160">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-160">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3a005-161">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-161">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="3a005-162">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-162">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-163">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-163">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="3a005-164">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-164">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="3a005-165">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-165">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="3a005-166">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="3a005-166">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="3a005-167">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-167">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="3a005-168">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-168">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="3a005-169">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="3a005-169">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="3a005-170">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="3a005-170">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="3a005-171">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-171">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="3a005-172">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-172">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="3a005-173">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-173">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="3a005-174">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-174">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-175">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="3a005-175">Office 2019 on Windows</span></span><br><span data-ttu-id="3a005-176">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3a005-176">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="3a005-177">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="3a005-177">
      - TaskPane</span></span><br><span data-ttu-id="3a005-178">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3a005-178">
      - Content</span></span><br><span data-ttu-id="3a005-179">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-179">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-180">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-180">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="3a005-181">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3a005-181">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="3a005-182">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3a005-182">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="3a005-183">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3a005-183">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="3a005-184">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3a005-184">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="3a005-185">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3a005-185">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="3a005-186">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3a005-186">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="3a005-187">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="3a005-187">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="3a005-188">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-188">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3a005-189">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-189">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-190">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-190">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="3a005-191">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-191">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="3a005-192">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-192">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="3a005-193">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="3a005-193">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="3a005-194">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-194">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="3a005-195">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-195">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="3a005-196">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="3a005-196">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="3a005-197">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="3a005-197">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="3a005-198">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-198">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="3a005-199">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-199">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="3a005-200">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-200">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="3a005-201">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-201">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-202">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="3a005-202">Office 2016 on Windows</span></span><br><span data-ttu-id="3a005-203">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3a005-203">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="3a005-204">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="3a005-204">
      - TaskPane</span></span><br><span data-ttu-id="3a005-205">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3a005-205">
      - Content</span></span> </td>
    <td><span data-ttu-id="3a005-206">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-206">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="3a005-207">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="3a005-207">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="3a005-208">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-208">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-209">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-209">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="3a005-210">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-210">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="3a005-211">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-211">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="3a005-212">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="3a005-212">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="3a005-213">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-213">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="3a005-214">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-214">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="3a005-215">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="3a005-215">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="3a005-216">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="3a005-216">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="3a005-217">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-217">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="3a005-218">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-218">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="3a005-219">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-219">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="3a005-220">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-220">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-221">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="3a005-221">Office 2013 on Windows</span></span><br><span data-ttu-id="3a005-222">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3a005-222">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="3a005-223">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="3a005-223">
      - TaskPane</span></span><br><span data-ttu-id="3a005-224">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3a005-224">
      - Content</span></span> </td>
    <td><span data-ttu-id="3a005-225">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="3a005-225">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="3a005-226">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-226">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-227">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-227">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="3a005-228">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-228">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="3a005-229">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="3a005-229">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="3a005-230">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-230">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="3a005-231">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-231">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="3a005-232">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="3a005-232">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="3a005-233">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="3a005-233">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="3a005-234">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-234">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="3a005-235">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-235">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="3a005-236">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-236">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="3a005-237">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-237">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-238">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="3a005-238">Office on iPad</span></span><br><span data-ttu-id="3a005-239">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="3a005-239">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="3a005-240">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="3a005-240">
      - TaskPane</span></span><br><span data-ttu-id="3a005-241">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3a005-241">
      - Content</span></span> </td>
    <td><span data-ttu-id="3a005-242">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-242">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="3a005-243">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3a005-243">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="3a005-244">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3a005-244">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="3a005-245">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3a005-245">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="3a005-246">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3a005-246">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="3a005-247">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3a005-247">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="3a005-248">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3a005-248">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="3a005-249">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="3a005-249">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="3a005-250">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="3a005-250">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="3a005-251">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="3a005-251">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="3a005-252">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="3a005-252">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="3a005-253">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-253">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3a005-254">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-254">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-255">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-255">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="3a005-256">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-256">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="3a005-257">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="3a005-257">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="3a005-258">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-258">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="3a005-259">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-259">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="3a005-260">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="3a005-260">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="3a005-261">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="3a005-261">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="3a005-262">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-262">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="3a005-263">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-263">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="3a005-264">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-264">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="3a005-265">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-265">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-266">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="3a005-266">Office on Mac</span></span><br><span data-ttu-id="3a005-267">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="3a005-267">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="3a005-268">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="3a005-268">
      - TaskPane</span></span><br><span data-ttu-id="3a005-269">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3a005-269">
      - Content</span></span><br><span data-ttu-id="3a005-270">
      - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="3a005-270">
      - Custom Functions</span></span><br><span data-ttu-id="3a005-271">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-271">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-272">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-272">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="3a005-273">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3a005-273">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="3a005-274">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3a005-274">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="3a005-275">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3a005-275">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="3a005-276">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3a005-276">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="3a005-277">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3a005-277">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="3a005-278">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3a005-278">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="3a005-279">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="3a005-279">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="3a005-280">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="3a005-280">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="3a005-281">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="3a005-281">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="3a005-282">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="3a005-282">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="3a005-283">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-283">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3a005-284">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-284">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="3a005-285">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-285">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-286">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-286">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="3a005-287">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-287">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="3a005-288">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-288">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="3a005-289">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="3a005-289">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="3a005-290">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-290">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="3a005-291">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-291">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="3a005-292">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-292">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="3a005-293">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="3a005-293">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="3a005-294">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="3a005-294">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="3a005-295">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-295">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="3a005-296">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-296">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="3a005-297">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-297">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="3a005-298">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-298">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-299">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="3a005-299">Office 2019 on Mac</span></span><br><span data-ttu-id="3a005-300">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3a005-300">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="3a005-301">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="3a005-301">
      - TaskPane</span></span><br><span data-ttu-id="3a005-302">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3a005-302">
      - Content</span></span><br><span data-ttu-id="3a005-303">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-303">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-304">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-304">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="3a005-305">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3a005-305">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="3a005-306">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3a005-306">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="3a005-307">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3a005-307">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="3a005-308">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3a005-308">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="3a005-309">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3a005-309">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="3a005-310">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3a005-310">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="3a005-311">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="3a005-311">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="3a005-312">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-312">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3a005-313">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-313">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-314">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-314">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="3a005-315">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-315">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="3a005-316">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-316">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="3a005-317">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="3a005-317">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="3a005-318">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-318">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="3a005-319">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-319">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="3a005-320">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-320">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="3a005-321">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="3a005-321">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="3a005-322">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="3a005-322">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="3a005-323">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-323">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="3a005-324">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-324">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="3a005-325">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-325">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="3a005-326">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-326">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-327">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="3a005-327">Office 2016 on Mac</span></span><br><span data-ttu-id="3a005-328">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3a005-328">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="3a005-329">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="3a005-329">
      - TaskPane</span></span><br><span data-ttu-id="3a005-330">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3a005-330">
      - Content</span></span> </td>
    <td><span data-ttu-id="3a005-331">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-331">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="3a005-332">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="3a005-332">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="3a005-333">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-333">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-334">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-334">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="3a005-335">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-335">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="3a005-336">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-336">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="3a005-337">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="3a005-337">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="3a005-338">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-338">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="3a005-339">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-339">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="3a005-340">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-340">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="3a005-341">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="3a005-341">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="3a005-342">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="3a005-342">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="3a005-343">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-343">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="3a005-344">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-344">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="3a005-345">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-345">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="3a005-346">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-346">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
</table>

<span data-ttu-id="3a005-347">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="3a005-347">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="3a005-348">Funções personalizadas (somente Excel)</span><span class="sxs-lookup"><span data-stu-id="3a005-348">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="3a005-349">Plataforma</span><span class="sxs-lookup"><span data-stu-id="3a005-349">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="3a005-350">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="3a005-350">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="3a005-351">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="3a005-351">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="3a005-352"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="3a005-352"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-353">Office na Web</span><span class="sxs-lookup"><span data-stu-id="3a005-353">Office on the web</span></span></td>
    <td><span data-ttu-id="3a005-354">- Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="3a005-354">- Custom Functions</span></span></td>
    <td><span data-ttu-id="3a005-355">- <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-355">- <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-356">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="3a005-356">Office on Windows</span></span><br><span data-ttu-id="3a005-357">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="3a005-357">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="3a005-358">- Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="3a005-358">- Custom Functions</span></span></td>
    <td><span data-ttu-id="3a005-359">- <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-359">- <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-360">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="3a005-360">Office on Mac</span></span><br><span data-ttu-id="3a005-361">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="3a005-361">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="3a005-362">- Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="3a005-362">- Custom Functions</span></span></td>
    <td><span data-ttu-id="3a005-363">- <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-363">- <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td></td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="3a005-364">Outlook</span><span class="sxs-lookup"><span data-stu-id="3a005-364">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="3a005-365">Plataforma</span><span class="sxs-lookup"><span data-stu-id="3a005-365">Platform</span></span></th>
    <th><span data-ttu-id="3a005-366">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="3a005-366">Extension points</span></span></th>
    <th><span data-ttu-id="3a005-367">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="3a005-367">API requirement sets</span></span></th>
    <th><span data-ttu-id="3a005-368"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="3a005-368"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-369">Office na Web</span><span class="sxs-lookup"><span data-stu-id="3a005-369">Office on the web</span></span><br><span data-ttu-id="3a005-370">(moderno)</span><span class="sxs-lookup"><span data-stu-id="3a005-370">(modern)</span></span></td>
    <td><span data-ttu-id="3a005-371">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="3a005-371">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="3a005-372">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="3a005-372">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="3a005-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="3a005-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="3a005-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="3a005-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="3a005-375">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-375">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3a005-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3a005-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3a005-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3a005-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3a005-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3a005-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="3a005-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3a005-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="3a005-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3a005-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="3a005-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3a005-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="3a005-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-384">Não disponível</span><span class="sxs-lookup"><span data-stu-id="3a005-384">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-385">Office na Web</span><span class="sxs-lookup"><span data-stu-id="3a005-385">Office on the web</span></span><br><span data-ttu-id="3a005-386">(clássico)</span><span class="sxs-lookup"><span data-stu-id="3a005-386">(classic)</span></span></td>
    <td><span data-ttu-id="3a005-387">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="3a005-387">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="3a005-388">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="3a005-388">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="3a005-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="3a005-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="3a005-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="3a005-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="3a005-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3a005-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3a005-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3a005-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3a005-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3a005-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3a005-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="3a005-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3a005-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="3a005-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-398">Não disponível</span><span class="sxs-lookup"><span data-stu-id="3a005-398">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-399">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="3a005-399">Office on Windows</span></span><br><span data-ttu-id="3a005-400">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="3a005-400">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="3a005-401">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="3a005-401">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="3a005-402">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="3a005-402">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="3a005-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="3a005-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="3a005-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="3a005-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="3a005-405">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="3a005-405">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="3a005-406">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Módulos</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-406">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3a005-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3a005-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3a005-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3a005-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3a005-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3a005-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="3a005-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3a005-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="3a005-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3a005-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="3a005-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3a005-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="3a005-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-415">Não disponível</span><span class="sxs-lookup"><span data-stu-id="3a005-415">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-416">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="3a005-416">Office 2019 on Windows</span></span><br><span data-ttu-id="3a005-417">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3a005-417">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="3a005-418">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="3a005-418">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="3a005-419">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="3a005-419">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="3a005-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="3a005-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="3a005-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="3a005-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="3a005-422">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="3a005-422">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="3a005-423">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Módulos</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-423">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3a005-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3a005-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3a005-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3a005-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3a005-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3a005-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="3a005-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3a005-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="3a005-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3a005-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="3a005-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-431">Não disponível</span><span class="sxs-lookup"><span data-stu-id="3a005-431">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-432">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="3a005-432">Office 2016 on Windows</span></span><br><span data-ttu-id="3a005-433">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3a005-433">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="3a005-434">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="3a005-434">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="3a005-435">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="3a005-435">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="3a005-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="3a005-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="3a005-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="3a005-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="3a005-438">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="3a005-438">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="3a005-439">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Módulos</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-439">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3a005-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3a005-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3a005-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3a005-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3a005-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*
    </span><span class="sxs-lookup"><span data-stu-id="3a005-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*
    </span></span></td>
    <td><span data-ttu-id="3a005-444">Não disponível</span><span class="sxs-lookup"><span data-stu-id="3a005-444">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-445">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="3a005-445">Office 2013 on Windows</span></span><br><span data-ttu-id="3a005-446">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3a005-446">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="3a005-447">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="3a005-447">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="3a005-448">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="3a005-448">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="3a005-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="3a005-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="3a005-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="3a005-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br>
    </td>
    <td><span data-ttu-id="3a005-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3a005-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3a005-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3a005-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="3a005-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="3a005-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*
    </span><span class="sxs-lookup"><span data-stu-id="3a005-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*
    </span></span></td>
    <td><span data-ttu-id="3a005-455">Não disponível</span><span class="sxs-lookup"><span data-stu-id="3a005-455">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-456">Office no iOS</span><span class="sxs-lookup"><span data-stu-id="3a005-456">Office on iOS</span></span><br><span data-ttu-id="3a005-457">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="3a005-457">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="3a005-458">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="3a005-458">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="3a005-459">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-459">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3a005-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3a005-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3a005-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3a005-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3a005-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3a005-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="3a005-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-465">Não disponível</span><span class="sxs-lookup"><span data-stu-id="3a005-465">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-466">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="3a005-466">Office on Mac</span></span><br><span data-ttu-id="3a005-467">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="3a005-467">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="3a005-468">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="3a005-468">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="3a005-469">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="3a005-469">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="3a005-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="3a005-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="3a005-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="3a005-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="3a005-472">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-472">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3a005-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3a005-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3a005-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3a005-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3a005-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3a005-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="3a005-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3a005-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="3a005-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3a005-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="3a005-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3a005-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="3a005-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-481">Não disponível</span><span class="sxs-lookup"><span data-stu-id="3a005-481">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-482">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="3a005-482">Office 2019 on Mac</span></span><br><span data-ttu-id="3a005-483">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3a005-483">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="3a005-484">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="3a005-484">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="3a005-485">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="3a005-485">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="3a005-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="3a005-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="3a005-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="3a005-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="3a005-488">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-488">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3a005-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3a005-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3a005-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3a005-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3a005-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3a005-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="3a005-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3a005-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="3a005-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-495">Não disponível</span><span class="sxs-lookup"><span data-stu-id="3a005-495">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-496">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="3a005-496">Office 2016 on Mac</span></span><br><span data-ttu-id="3a005-497">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3a005-497">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="3a005-498">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="3a005-498">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="3a005-499">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="3a005-499">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="3a005-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="3a005-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="3a005-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="3a005-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="3a005-502">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-502">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-503">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-503">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3a005-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3a005-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3a005-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3a005-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3a005-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3a005-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="3a005-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3a005-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="3a005-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-509">Não disponível</span><span class="sxs-lookup"><span data-stu-id="3a005-509">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-510">Outlook no Android</span><span class="sxs-lookup"><span data-stu-id="3a005-510">Office on Android</span></span><br><span data-ttu-id="3a005-511">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="3a005-511">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="3a005-512">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="3a005-512">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="3a005-513">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Organizador de compromissos (Redigir): reunião on-line (visualização)</span><span class="sxs-lookup"><span data-stu-id="3a005-513">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Appointment Organizer (Compose): online meeting</a> (preview)</span></span><br><span data-ttu-id="3a005-514">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-514">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-515">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-515">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3a005-516">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3a005-516">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3a005-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3a005-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3a005-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3a005-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="3a005-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-520">Não disponível</span><span class="sxs-lookup"><span data-stu-id="3a005-520">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="3a005-521">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="3a005-521">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3a005-522">O suporte ao cliente para um conjunto de requisitos pode ser restringido pelo suporte do servidor Exchange.</span><span class="sxs-lookup"><span data-stu-id="3a005-522">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="3a005-523">Consulte [Conjuntos de requisitos da API JavaScript do Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) para obter detalhes sobre o intervalo de conjuntos de requisitos suportado pelo servidor Exchange e pelos clientes Outlook.</span><span class="sxs-lookup"><span data-stu-id="3a005-523">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="3a005-524">Word</span><span class="sxs-lookup"><span data-stu-id="3a005-524">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="3a005-525">Plataforma</span><span class="sxs-lookup"><span data-stu-id="3a005-525">Platform</span></span></th>
    <th><span data-ttu-id="3a005-526">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="3a005-526">Extension points</span></span></th>
    <th><span data-ttu-id="3a005-527">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="3a005-527">API requirement sets</span></span></th>
    <th><span data-ttu-id="3a005-528"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="3a005-528"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-529">Office na Web</span><span class="sxs-lookup"><span data-stu-id="3a005-529">Office on the web</span></span></td>
    <td><span data-ttu-id="3a005-530">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="3a005-530">
      - TaskPane</span></span><br><span data-ttu-id="3a005-531">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-531">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-532">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-532">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="3a005-533">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3a005-533">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="3a005-534">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3a005-534">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="3a005-535">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-535">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3a005-536">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-536">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="3a005-537">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-537">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-538">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-538">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="3a005-539">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="3a005-539">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="3a005-540">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-540">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="3a005-541">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="3a005-541">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="3a005-542">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-542">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="3a005-543">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-543">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="3a005-544">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-544">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="3a005-545">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-545">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="3a005-546">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-546">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="3a005-547">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="3a005-547">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="3a005-548">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="3a005-548">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="3a005-549">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-549">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="3a005-550">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-550">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="3a005-551">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-551">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="3a005-552">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-552">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="3a005-553">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-553">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-554">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="3a005-554">Office on Windows</span></span><br><span data-ttu-id="3a005-555">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="3a005-555">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="3a005-556">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="3a005-556">
      - TaskPane</span></span><br><span data-ttu-id="3a005-557">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-557">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-558">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-558">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="3a005-559">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3a005-559">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="3a005-560">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3a005-560">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="3a005-561">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-561">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3a005-562">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-562">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="3a005-563">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-563">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-564">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-564">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="3a005-565">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-565">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="3a005-566">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="3a005-566">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="3a005-567">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-567">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="3a005-568">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="3a005-568">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="3a005-569">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-569">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="3a005-570">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-570">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="3a005-571">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-571">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="3a005-572">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-572">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="3a005-573">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-573">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="3a005-574">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="3a005-574">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="3a005-575">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="3a005-575">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="3a005-576">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-576">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="3a005-577">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-577">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="3a005-578">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-578">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="3a005-579">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-579">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="3a005-580">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-580">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-581">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="3a005-581">Office 2019 on Windows</span></span><br><span data-ttu-id="3a005-582">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3a005-582">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="3a005-583">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="3a005-583">
      - TaskPane</span></span><br><span data-ttu-id="3a005-584">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-584">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-585">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-585">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="3a005-586">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3a005-586">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="3a005-587">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3a005-587">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="3a005-588">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-588">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3a005-589">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-589">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-590">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-590">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="3a005-591">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-591">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="3a005-592">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="3a005-592">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="3a005-593">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-593">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="3a005-594">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="3a005-594">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="3a005-595">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-595">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="3a005-596">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-596">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="3a005-597">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-597">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="3a005-598">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-598">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="3a005-599">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-599">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="3a005-600">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="3a005-600">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="3a005-601">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="3a005-601">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="3a005-602">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-602">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="3a005-603">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-603">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="3a005-604">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-604">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="3a005-605">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-605">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="3a005-606">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-606">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-607">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="3a005-607">Office 2016 on Windows</span></span><br><span data-ttu-id="3a005-608">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3a005-608">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="3a005-609">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="3a005-609">- TaskPane</span></span></td>
    <td><span data-ttu-id="3a005-610">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-610">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="3a005-611">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="3a005-611">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="3a005-612">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-612">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-613">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-613">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="3a005-614">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-614">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="3a005-615">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="3a005-615">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="3a005-616">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-616">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="3a005-617">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="3a005-617">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="3a005-618">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-618">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="3a005-619">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-619">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="3a005-620">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-620">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="3a005-621">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-621">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="3a005-622">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-622">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="3a005-623">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="3a005-623">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="3a005-624">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="3a005-624">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="3a005-625">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-625">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="3a005-626">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-626">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="3a005-627">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-627">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="3a005-628">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-628">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="3a005-629">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-629">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-630">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="3a005-630">Office 2013 on Windows</span></span><br><span data-ttu-id="3a005-631">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3a005-631">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="3a005-632">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="3a005-632">- TaskPane</span></span></td>
    <td><span data-ttu-id="3a005-633">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="3a005-633">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="3a005-634">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-634">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-635">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-635">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="3a005-636">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-636">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="3a005-637">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="3a005-637">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="3a005-638">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-638">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="3a005-639">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="3a005-639">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="3a005-640">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-640">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="3a005-641">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-641">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="3a005-642">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-642">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="3a005-643">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-643">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="3a005-644">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-644">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="3a005-645">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="3a005-645">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="3a005-646">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="3a005-646">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="3a005-647">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-647">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="3a005-648">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-648">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="3a005-649">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-649">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="3a005-650">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-650">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="3a005-651">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-651">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-652">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="3a005-652">Office on iPad</span></span><br><span data-ttu-id="3a005-653">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="3a005-653">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="3a005-654">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="3a005-654">- TaskPane</span></span></td>
    <td><span data-ttu-id="3a005-655">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-655">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="3a005-656">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3a005-656">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="3a005-657">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3a005-657">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="3a005-658">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-658">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3a005-659">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-659">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-660">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-660">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="3a005-661">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-661">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="3a005-662">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="3a005-662">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="3a005-663">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-663">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="3a005-664">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="3a005-664">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="3a005-665">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-665">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="3a005-666">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-666">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="3a005-667">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-667">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="3a005-668">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-668">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="3a005-669">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-669">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="3a005-670">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="3a005-670">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="3a005-671">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="3a005-671">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="3a005-672">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-672">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="3a005-673">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-673">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="3a005-674">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-674">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="3a005-675">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-675">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="3a005-676">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-676">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-677">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="3a005-677">Office on Mac</span></span><br><span data-ttu-id="3a005-678">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="3a005-678">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="3a005-679">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="3a005-679">
      - TaskPane</span></span><br><span data-ttu-id="3a005-680">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-680">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-681">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-681">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="3a005-682">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3a005-682">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="3a005-683">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3a005-683">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="3a005-684">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-684">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3a005-685">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-685">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="3a005-686">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-686">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-687">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-687">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="3a005-688">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-688">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="3a005-689">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="3a005-689">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="3a005-690">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-690">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="3a005-691">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="3a005-691">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="3a005-692">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-692">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="3a005-693">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-693">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="3a005-694">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-694">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="3a005-695">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-695">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="3a005-696">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-696">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="3a005-697">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="3a005-697">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="3a005-698">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="3a005-698">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="3a005-699">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-699">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="3a005-700">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-700">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="3a005-701">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-701">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="3a005-702">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-702">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="3a005-703">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-703">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-704">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="3a005-704">Office 2019 on Mac</span></span><br><span data-ttu-id="3a005-705">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3a005-705">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="3a005-706">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="3a005-706">
      - TaskPane</span></span><br><span data-ttu-id="3a005-707">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-707">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-708">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-708">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="3a005-709">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3a005-709">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="3a005-710">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3a005-710">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="3a005-711">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-711">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3a005-712">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-712">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-713">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-713">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="3a005-714">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-714">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="3a005-715">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="3a005-715">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="3a005-716">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-716">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="3a005-717">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="3a005-717">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="3a005-718">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-718">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="3a005-719">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-719">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="3a005-720">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-720">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="3a005-721">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-721">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="3a005-722">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-722">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="3a005-723">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="3a005-723">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="3a005-724">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="3a005-724">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="3a005-725">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-725">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="3a005-726">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-726">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="3a005-727">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-727">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="3a005-728">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-728">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="3a005-729">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-729">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-730">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="3a005-730">Office 2016 on Mac</span></span><br><span data-ttu-id="3a005-731">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3a005-731">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="3a005-732">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="3a005-732">- TaskPane</span></span></td>
    <td><span data-ttu-id="3a005-733">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-733">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="3a005-734">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="3a005-734">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="3a005-735">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-735">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-736">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-736">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="3a005-737">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-737">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="3a005-738">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="3a005-738">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="3a005-739">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-739">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="3a005-740">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="3a005-740">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="3a005-741">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-741">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="3a005-742">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-742">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="3a005-743">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-743">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="3a005-744">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-744">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="3a005-745">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-745">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="3a005-746">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="3a005-746">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="3a005-747">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="3a005-747">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="3a005-748">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-748">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="3a005-749">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-749">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="3a005-750">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="3a005-750">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="3a005-751">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-751">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="3a005-752">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-752">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
</table>

<span data-ttu-id="3a005-753">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="3a005-753">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="3a005-754">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="3a005-754">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="3a005-755">Plataforma</span><span class="sxs-lookup"><span data-stu-id="3a005-755">Platform</span></span></th>
    <th><span data-ttu-id="3a005-756">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="3a005-756">Extension points</span></span></th>
    <th><span data-ttu-id="3a005-757">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="3a005-757">API requirement sets</span></span></th>
    <th><span data-ttu-id="3a005-758"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="3a005-758"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-759">Office na Web</span><span class="sxs-lookup"><span data-stu-id="3a005-759">Office on the web</span></span></td>
    <td><span data-ttu-id="3a005-760">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3a005-760">
      - Content</span></span><br><span data-ttu-id="3a005-761">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="3a005-761">
      - TaskPane</span></span><br><span data-ttu-id="3a005-762">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-762">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-763">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-763">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="3a005-764">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-764">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3a005-765">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-765">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="3a005-766">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-766">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-767">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="3a005-767">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="3a005-768">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-768">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="3a005-769">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-769">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="3a005-770">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="3a005-770">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="3a005-771">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-771">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="3a005-772">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="3a005-772">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="3a005-773">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="3a005-773">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="3a005-774">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-774">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-775">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="3a005-775">Office on Windows</span></span><br><span data-ttu-id="3a005-776">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="3a005-776">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="3a005-777">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3a005-777">
      - Content</span></span><br><span data-ttu-id="3a005-778">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="3a005-778">
      - TaskPane</span></span><br><span data-ttu-id="3a005-779">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-779">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-780">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-780">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="3a005-781">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-781">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3a005-782">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-782">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="3a005-783">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-783">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-784">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="3a005-784">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="3a005-785">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-785">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="3a005-786">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-786">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="3a005-787">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="3a005-787">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="3a005-788">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-788">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="3a005-789">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="3a005-789">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="3a005-790">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="3a005-790">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="3a005-791">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-791">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-792">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="3a005-792">Office 2019 on Windows</span></span><br><span data-ttu-id="3a005-793">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3a005-793">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="3a005-794">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3a005-794">
      - Content</span></span><br><span data-ttu-id="3a005-795">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="3a005-795">
      - TaskPane</span></span><br><span data-ttu-id="3a005-796">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-796">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-797">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-797">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3a005-798">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-798">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-799">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="3a005-799">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="3a005-800">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-800">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="3a005-801">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-801">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="3a005-802">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="3a005-802">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="3a005-803">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-803">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="3a005-804">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="3a005-804">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="3a005-805">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="3a005-805">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="3a005-806">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-806">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-807">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="3a005-807">Office 2016 on Windows</span></span><br><span data-ttu-id="3a005-808">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3a005-808">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="3a005-809">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3a005-809">
      - Content</span></span><br><span data-ttu-id="3a005-810">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="3a005-810">
      - TaskPane</span></span> </td>
    <td><span data-ttu-id="3a005-811">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="3a005-811">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="3a005-812">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-812">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-813">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="3a005-813">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="3a005-814">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-814">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="3a005-815">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-815">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="3a005-816">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="3a005-816">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="3a005-817">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-817">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="3a005-818">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="3a005-818">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="3a005-819">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="3a005-819">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="3a005-820">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-820">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-821">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="3a005-821">Office 2013 on Windows</span></span><br><span data-ttu-id="3a005-822">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3a005-822">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="3a005-823">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3a005-823">
      - Content</span></span><br><span data-ttu-id="3a005-824">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="3a005-824">
      - TaskPane</span></span> </td>
    <td><span data-ttu-id="3a005-825">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="3a005-825">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="3a005-826">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-826">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-827">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="3a005-827">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="3a005-828">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-828">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="3a005-829">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-829">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="3a005-830">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="3a005-830">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="3a005-831">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-831">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="3a005-832">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="3a005-832">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="3a005-833">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="3a005-833">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="3a005-834">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-834">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-835">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="3a005-835">Office on iPad</span></span><br><span data-ttu-id="3a005-836">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="3a005-836">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="3a005-837">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3a005-837">
      - Content</span></span><br><span data-ttu-id="3a005-838">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="3a005-838">
      - TaskPane</span></span> </td>
    <td><span data-ttu-id="3a005-839">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-839">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="3a005-840">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-840">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3a005-841">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-841">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-842">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="3a005-842">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="3a005-843">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-843">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="3a005-844">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-844">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="3a005-845">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="3a005-845">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="3a005-846">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-846">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="3a005-847">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="3a005-847">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="3a005-848">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="3a005-848">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="3a005-849">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-849">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-850">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="3a005-850">Office on Mac</span></span><br><span data-ttu-id="3a005-851">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="3a005-851">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="3a005-852">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3a005-852">
      - Content</span></span><br><span data-ttu-id="3a005-853">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="3a005-853">
      - TaskPane</span></span><br><span data-ttu-id="3a005-854">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-854">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-855">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-855">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="3a005-856">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-856">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3a005-857">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-857">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="3a005-858">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-858">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-859">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="3a005-859">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="3a005-860">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-860">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="3a005-861">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-861">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="3a005-862">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="3a005-862">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="3a005-863">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-863">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="3a005-864">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="3a005-864">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="3a005-865">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="3a005-865">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="3a005-866">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-866">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-867">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="3a005-867">Office 2019 on Mac</span></span><br><span data-ttu-id="3a005-868">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3a005-868">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="3a005-869">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3a005-869">
      - Content</span></span><br><span data-ttu-id="3a005-870">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="3a005-870">
      - TaskPane</span></span><br><span data-ttu-id="3a005-871">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-871">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-872">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-872">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3a005-873">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-873">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-874">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="3a005-874">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="3a005-875">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-875">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="3a005-876">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-876">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="3a005-877">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="3a005-877">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="3a005-878">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-878">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="3a005-879">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="3a005-879">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="3a005-880">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="3a005-880">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="3a005-881">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-881">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-882">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="3a005-882">Office 2016 on Mac</span></span><br><span data-ttu-id="3a005-883">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3a005-883">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="3a005-884">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3a005-884">
      - Content</span></span><br><span data-ttu-id="3a005-885">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="3a005-885">
      - TaskPane</span></span> </td>
    <td><span data-ttu-id="3a005-886">
       - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="3a005-886">
       - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="3a005-887">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-887">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-888">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="3a005-888">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="3a005-889">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-889">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="3a005-890">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-890">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="3a005-891">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="3a005-891">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="3a005-892">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="3a005-892">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="3a005-893">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="3a005-893">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="3a005-894">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="3a005-894">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="3a005-895">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-895">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
</table>

<span data-ttu-id="3a005-896">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="3a005-896">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="3a005-897">OneNote</span><span class="sxs-lookup"><span data-stu-id="3a005-897">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="3a005-898">Plataforma</span><span class="sxs-lookup"><span data-stu-id="3a005-898">Platform</span></span></th>
    <th><span data-ttu-id="3a005-899">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="3a005-899">Extension points</span></span></th>
    <th><span data-ttu-id="3a005-900">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="3a005-900">API requirement sets</span></span></th>
    <th><span data-ttu-id="3a005-901"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="3a005-901"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-902">Office na Web</span><span class="sxs-lookup"><span data-stu-id="3a005-902">Office on the web</span></span></td>
    <td><span data-ttu-id="3a005-903">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3a005-903">
      - Content</span></span><br><span data-ttu-id="3a005-904">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="3a005-904">
      - TaskPane</span></span><br><span data-ttu-id="3a005-905">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-905">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-906">
      - <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-906">
      - <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="3a005-907">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-907">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3a005-908">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-908">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="3a005-909">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="3a005-909">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="3a005-910">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="3a005-910">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="3a005-911">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="3a005-911">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="3a005-912">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-912">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="3a005-913">Project</span><span class="sxs-lookup"><span data-stu-id="3a005-913">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="3a005-914">Plataforma</span><span class="sxs-lookup"><span data-stu-id="3a005-914">Platform</span></span></th>
    <th><span data-ttu-id="3a005-915">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="3a005-915">Extension points</span></span></th>
    <th><span data-ttu-id="3a005-916">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="3a005-916">API requirement sets</span></span></th>
    <th><span data-ttu-id="3a005-917"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="3a005-917"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-918">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="3a005-918">Office 2019 on Windows</span></span><br><span data-ttu-id="3a005-919">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3a005-919">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="3a005-920">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="3a005-920">- TaskPane</span></span></td>
    <td><span data-ttu-id="3a005-921">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-921">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="3a005-922">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="3a005-922">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="3a005-923">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-923">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-924">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="3a005-924">Office 2016 on Windows</span></span><br><span data-ttu-id="3a005-925">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3a005-925">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="3a005-926">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="3a005-926">- TaskPane</span></span></td>
    <td><span data-ttu-id="3a005-927">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-927">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="3a005-928">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="3a005-928">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="3a005-929">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-929">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3a005-930">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="3a005-930">Office 2013 on Windows</span></span><br><span data-ttu-id="3a005-931">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3a005-931">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="3a005-932">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="3a005-932">- TaskPane</span></span></td>
    <td><span data-ttu-id="3a005-933">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3a005-933">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="3a005-934">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="3a005-934">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="3a005-935">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="3a005-935">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="3a005-936">Confira também</span><span class="sxs-lookup"><span data-stu-id="3a005-936">See also</span></span>

- [<span data-ttu-id="3a005-937">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="3a005-937">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="3a005-938">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="3a005-938">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="3a005-939">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="3a005-939">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="3a005-940">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="3a005-940">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="3a005-941">Documentação de Referência da API</span><span class="sxs-lookup"><span data-stu-id="3a005-941">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="3a005-942">Histórico de atualizações para Microsoft 365 Apps</span><span class="sxs-lookup"><span data-stu-id="3a005-942">Update history for Microsoft 365 Apps</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="3a005-943">Histórico de atualizações do Office 2016 e 2019 (Clique para Executar)</span><span class="sxs-lookup"><span data-stu-id="3a005-943">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="3a005-944">Histórico de atualizações do Office 2013 (clique para executar)</span><span class="sxs-lookup"><span data-stu-id="3a005-944">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="3a005-945">Histórico de atualizações do Office 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="3a005-945">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="3a005-946">Histórico de atualizações do Outlook 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="3a005-946">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="3a005-947">Histórico de atualizações do Office para Mac</span><span class="sxs-lookup"><span data-stu-id="3a005-947">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="3a005-948">Criando Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="3a005-948">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
