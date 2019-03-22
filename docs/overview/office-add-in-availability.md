---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos compatíveis com Excel, Word, Outlook, PowerPoint, OneNote e Project.
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: fe5b1d1278d2c14192fb6fd212f24bb08571d35d
ms.sourcegitcommit: c5daedf017c6dd5ab0c13607589208c3f3627354
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/20/2019
ms.locfileid: "30691122"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="4af34-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="4af34-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="4af34-104">Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="4af34-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="4af34-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que são compatíveis atualmente com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="4af34-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="4af34-p102">O número do build do Office 2016 instalado via MSI é 16.0.4266.1001. Esta versão só contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="4af34-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span>
>
> <span data-ttu-id="4af34-108">O número de build para uma compra avulsa do Office 2019 é 16.0.10827.20150.</span><span class="sxs-lookup"><span data-stu-id="4af34-108">The build number for a one-time purchase of Office 2019 is 16.0.10827.20150.</span></span>

## <a name="excel"></a><span data-ttu-id="4af34-109">Excel</span><span class="sxs-lookup"><span data-stu-id="4af34-109">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="4af34-110">Plataforma</span><span class="sxs-lookup"><span data-stu-id="4af34-110">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="4af34-111">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="4af34-111">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="4af34-112">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="4af34-112">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="4af34-113"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="4af34-113"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-114">Office Online</span><span class="sxs-lookup"><span data-stu-id="4af34-114">Office Online</span></span></td>
    <td> <span data-ttu-id="4af34-115">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="4af34-115">- TaskPane</span></span><br><span data-ttu-id="4af34-116">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="4af34-116">
        - Content</span></span><br><span data-ttu-id="4af34-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="4af34-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4af34-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4af34-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4af34-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4af34-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4af34-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4af34-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4af34-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4af34-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4af34-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4af34-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4af34-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4af34-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4af34-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4af34-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4af34-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4af34-126">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-126">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4af34-127">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-127">
        - BindingEvents</span></span><br><span data-ttu-id="4af34-128">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4af34-128">
        - CompressedFile</span></span><br><span data-ttu-id="4af34-129">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-129">
        - DocumentEvents</span></span><br><span data-ttu-id="4af34-130">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4af34-130">
        - File</span></span><br><span data-ttu-id="4af34-131">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-131">
        - MatrixBindings</span></span><br><span data-ttu-id="4af34-132">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-132">
        - MatrixCoercion</span></span><br><span data-ttu-id="4af34-133">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="4af34-133">
        - Selection</span></span><br><span data-ttu-id="4af34-134">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="4af34-134">
        - Settings</span></span><br><span data-ttu-id="4af34-135">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-135">
        - TableBindings</span></span><br><span data-ttu-id="4af34-136">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-136">
        - TableCoercion</span></span><br><span data-ttu-id="4af34-137">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-137">
        - TextBindings</span></span><br><span data-ttu-id="4af34-138">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-138">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-139">Office 365 para Windows</span><span class="sxs-lookup"><span data-stu-id="4af34-139">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="4af34-140">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="4af34-140">- TaskPane</span></span><br><span data-ttu-id="4af34-141">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="4af34-141">
        - Content</span></span><br><span data-ttu-id="4af34-142">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="4af34-142">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4af34-143">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-143">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4af34-144">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4af34-144">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4af34-145">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4af34-145">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4af34-146">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4af34-146">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4af34-147">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4af34-147">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4af34-148">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4af34-148">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4af34-149">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4af34-149">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4af34-150">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4af34-150">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4af34-151">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-151">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4af34-152">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-152">
        - BindingEvents</span></span><br><span data-ttu-id="4af34-153">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4af34-153">
        - CompressedFile</span></span><br><span data-ttu-id="4af34-154">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-154">
        - DocumentEvents</span></span><br><span data-ttu-id="4af34-155">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4af34-155">
        - File</span></span><br><span data-ttu-id="4af34-156">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-156">
        - MatrixBindings</span></span><br><span data-ttu-id="4af34-157">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-157">
        - MatrixCoercion</span></span><br><span data-ttu-id="4af34-158">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="4af34-158">
        - Selection</span></span><br><span data-ttu-id="4af34-159">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="4af34-159">
        - Settings</span></span><br><span data-ttu-id="4af34-160">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-160">
        - TableBindings</span></span><br><span data-ttu-id="4af34-161">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-161">
        - TableCoercion</span></span><br><span data-ttu-id="4af34-162">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-162">
        - TextBindings</span></span><br><span data-ttu-id="4af34-163">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-163">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-164">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="4af34-164">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="4af34-165">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="4af34-165">- TaskPane</span></span><br><span data-ttu-id="4af34-166">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="4af34-166">
        - Content</span></span><br><span data-ttu-id="4af34-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4af34-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="4af34-168">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-168">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4af34-169">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4af34-169">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4af34-170">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4af34-170">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4af34-171">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4af34-171">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4af34-172">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4af34-172">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4af34-173">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4af34-173">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4af34-174">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4af34-174">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4af34-175">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4af34-175">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4af34-176">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-176">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4af34-177">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-177">- BindingEvents</span></span><br><span data-ttu-id="4af34-178">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4af34-178">
        - CompressedFile</span></span><br><span data-ttu-id="4af34-179">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-179">
        - DocumentEvents</span></span><br><span data-ttu-id="4af34-180">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4af34-180">
        - File</span></span><br><span data-ttu-id="4af34-181">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-181">
        - ImageCoercion</span></span><br><span data-ttu-id="4af34-182">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-182">
        - MatrixBindings</span></span><br><span data-ttu-id="4af34-183">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-183">
        - MatrixCoercion</span></span><br><span data-ttu-id="4af34-184">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="4af34-184">
        - Selection</span></span><br><span data-ttu-id="4af34-185">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="4af34-185">
        - Settings</span></span><br><span data-ttu-id="4af34-186">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-186">
        - TableBindings</span></span><br><span data-ttu-id="4af34-187">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-187">
        - TableCoercion</span></span><br><span data-ttu-id="4af34-188">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-188">
        - TextBindings</span></span><br><span data-ttu-id="4af34-189">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-189">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-190">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="4af34-190">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="4af34-191">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="4af34-191">- TaskPane</span></span><br><span data-ttu-id="4af34-192">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="4af34-192">
        - Content</span></span></td>
    <td><span data-ttu-id="4af34-193">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-193">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4af34-194">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="4af34-194">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="4af34-195">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-195">- BindingEvents</span></span><br><span data-ttu-id="4af34-196">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4af34-196">
        - CompressedFile</span></span><br><span data-ttu-id="4af34-197">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-197">
        - DocumentEvents</span></span><br><span data-ttu-id="4af34-198">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4af34-198">
        - File</span></span><br><span data-ttu-id="4af34-199">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-199">
        - ImageCoercion</span></span><br><span data-ttu-id="4af34-200">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-200">
        - MatrixBindings</span></span><br><span data-ttu-id="4af34-201">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-201">
        - MatrixCoercion</span></span><br><span data-ttu-id="4af34-202">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="4af34-202">
        - Selection</span></span><br><span data-ttu-id="4af34-203">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="4af34-203">
        - Settings</span></span><br><span data-ttu-id="4af34-204">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-204">
        - TableBindings</span></span><br><span data-ttu-id="4af34-205">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-205">
        - TableCoercion</span></span><br><span data-ttu-id="4af34-206">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-206">
        - TextBindings</span></span><br><span data-ttu-id="4af34-207">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-207">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-208">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="4af34-208">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="4af34-209">
        - TaskPane</span><span class="sxs-lookup"><span data-stu-id="4af34-209">
        - TaskPane</span></span><br><span data-ttu-id="4af34-210">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="4af34-210">
        - Content</span></span></td>
    <td>  <span data-ttu-id="4af34-211">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="4af34-211">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="4af34-212">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-212">
        - BindingEvents</span></span><br><span data-ttu-id="4af34-213">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4af34-213">
        - CompressedFile</span></span><br><span data-ttu-id="4af34-214">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-214">
        - DocumentEvents</span></span><br><span data-ttu-id="4af34-215">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4af34-215">
        - File</span></span><br><span data-ttu-id="4af34-216">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-216">
        - ImageCoercion</span></span><br><span data-ttu-id="4af34-217">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-217">
        - MatrixBindings</span></span><br><span data-ttu-id="4af34-218">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-218">
        - MatrixCoercion</span></span><br><span data-ttu-id="4af34-219">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="4af34-219">
        - Selection</span></span><br><span data-ttu-id="4af34-220">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="4af34-220">
        - Settings</span></span><br><span data-ttu-id="4af34-221">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-221">
        - TableBindings</span></span><br><span data-ttu-id="4af34-222">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-222">
        - TableCoercion</span></span><br><span data-ttu-id="4af34-223">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-223">
        - TextBindings</span></span><br><span data-ttu-id="4af34-224">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-224">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-225">Office 365 para iPad</span><span class="sxs-lookup"><span data-stu-id="4af34-225">Office 365 for iPad</span></span></td>
    <td><span data-ttu-id="4af34-226">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="4af34-226">- TaskPane</span></span><br><span data-ttu-id="4af34-227">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="4af34-227">
        - Content</span></span></td>
    <td><span data-ttu-id="4af34-228">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-228">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4af34-229">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4af34-229">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4af34-230">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4af34-230">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4af34-231">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4af34-231">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4af34-232">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4af34-232">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4af34-233">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4af34-233">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4af34-234">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4af34-234">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4af34-235">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4af34-235">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4af34-236">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-236">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4af34-237">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-237">- BindingEvents</span></span><br><span data-ttu-id="4af34-238">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4af34-238">
        - CompressedFile</span></span><br><span data-ttu-id="4af34-239">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-239">
        - DocumentEvents</span></span><br><span data-ttu-id="4af34-240">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4af34-240">
        - File</span></span><br><span data-ttu-id="4af34-241">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-241">
        - ImageCoercion</span></span><br><span data-ttu-id="4af34-242">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-242">
        - MatrixBindings</span></span><br><span data-ttu-id="4af34-243">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-243">
        - MatrixCoercion</span></span><br><span data-ttu-id="4af34-244">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="4af34-244">
        - Selection</span></span><br><span data-ttu-id="4af34-245">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="4af34-245">
        - Settings</span></span><br><span data-ttu-id="4af34-246">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-246">
        - TableBindings</span></span><br><span data-ttu-id="4af34-247">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-247">
        - TableCoercion</span></span><br><span data-ttu-id="4af34-248">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-248">
        - TextBindings</span></span><br><span data-ttu-id="4af34-249">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-249">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-250">Office 365 para Mac</span><span class="sxs-lookup"><span data-stu-id="4af34-250">Office 365 for Mac</span></span></td>
    <td><span data-ttu-id="4af34-251">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="4af34-251">- TaskPane</span></span><br><span data-ttu-id="4af34-252">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="4af34-252">
        - Content</span></span><br><span data-ttu-id="4af34-253">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4af34-253">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="4af34-254">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-254">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4af34-255">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4af34-255">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4af34-256">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4af34-256">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4af34-257">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4af34-257">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4af34-258">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4af34-258">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4af34-259">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4af34-259">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4af34-260">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4af34-260">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4af34-261">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4af34-261">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4af34-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4af34-263">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-263">- BindingEvents</span></span><br><span data-ttu-id="4af34-264">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4af34-264">
        - CompressedFile</span></span><br><span data-ttu-id="4af34-265">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-265">
        - DocumentEvents</span></span><br><span data-ttu-id="4af34-266">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4af34-266">
        - File</span></span><br><span data-ttu-id="4af34-267">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-267">
        - ImageCoercion</span></span><br><span data-ttu-id="4af34-268">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-268">
        - MatrixBindings</span></span><br><span data-ttu-id="4af34-269">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-269">
        - MatrixCoercion</span></span><br><span data-ttu-id="4af34-270">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4af34-270">
        - PdfFile</span></span><br><span data-ttu-id="4af34-271">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="4af34-271">
        - Selection</span></span><br><span data-ttu-id="4af34-272">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="4af34-272">
        - Settings</span></span><br><span data-ttu-id="4af34-273">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-273">
        - TableBindings</span></span><br><span data-ttu-id="4af34-274">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-274">
        - TableCoercion</span></span><br><span data-ttu-id="4af34-275">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-275">
        - TextBindings</span></span><br><span data-ttu-id="4af34-276">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-276">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-277">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="4af34-277">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="4af34-278">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="4af34-278">- TaskPane</span></span><br><span data-ttu-id="4af34-279">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="4af34-279">
        - Content</span></span><br><span data-ttu-id="4af34-280">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4af34-280">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="4af34-281">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-281">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4af34-282">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4af34-282">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4af34-283">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4af34-283">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4af34-284">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4af34-284">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4af34-285">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4af34-285">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4af34-286">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4af34-286">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4af34-287">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4af34-287">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4af34-288">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4af34-288">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4af34-289">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-289">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4af34-290">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-290">- BindingEvents</span></span><br><span data-ttu-id="4af34-291">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4af34-291">
        - CompressedFile</span></span><br><span data-ttu-id="4af34-292">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-292">
        - DocumentEvents</span></span><br><span data-ttu-id="4af34-293">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4af34-293">
        - File</span></span><br><span data-ttu-id="4af34-294">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-294">
        - ImageCoercion</span></span><br><span data-ttu-id="4af34-295">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-295">
        - MatrixBindings</span></span><br><span data-ttu-id="4af34-296">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-296">
        - MatrixCoercion</span></span><br><span data-ttu-id="4af34-297">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4af34-297">
        - PdfFile</span></span><br><span data-ttu-id="4af34-298">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="4af34-298">
        - Selection</span></span><br><span data-ttu-id="4af34-299">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="4af34-299">
        - Settings</span></span><br><span data-ttu-id="4af34-300">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-300">
        - TableBindings</span></span><br><span data-ttu-id="4af34-301">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-301">
        - TableCoercion</span></span><br><span data-ttu-id="4af34-302">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-302">
        - TextBindings</span></span><br><span data-ttu-id="4af34-303">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-303">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-304">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="4af34-304">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="4af34-305">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="4af34-305">- TaskPane</span></span><br><span data-ttu-id="4af34-306">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="4af34-306">
        - Content</span></span></td>
    <td><span data-ttu-id="4af34-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4af34-308">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="4af34-308">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="4af34-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-309">- BindingEvents</span></span><br><span data-ttu-id="4af34-310">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4af34-310">
        - CompressedFile</span></span><br><span data-ttu-id="4af34-311">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-311">
        - DocumentEvents</span></span><br><span data-ttu-id="4af34-312">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4af34-312">
        - File</span></span><br><span data-ttu-id="4af34-313">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-313">
        - ImageCoercion</span></span><br><span data-ttu-id="4af34-314">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-314">
        - MatrixBindings</span></span><br><span data-ttu-id="4af34-315">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-315">
        - MatrixCoercion</span></span><br><span data-ttu-id="4af34-316">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4af34-316">
        - PdfFile</span></span><br><span data-ttu-id="4af34-317">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="4af34-317">
        - Selection</span></span><br><span data-ttu-id="4af34-318">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="4af34-318">
        - Settings</span></span><br><span data-ttu-id="4af34-319">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-319">
        - TableBindings</span></span><br><span data-ttu-id="4af34-320">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-320">
        - TableCoercion</span></span><br><span data-ttu-id="4af34-321">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-321">
        - TextBindings</span></span><br><span data-ttu-id="4af34-322">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-322">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="4af34-323">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="4af34-323">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="outlook"></a><span data-ttu-id="4af34-324">Outlook</span><span class="sxs-lookup"><span data-stu-id="4af34-324">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4af34-325">Plataforma</span><span class="sxs-lookup"><span data-stu-id="4af34-325">Platform</span></span></th>
    <th><span data-ttu-id="4af34-326">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="4af34-326">Extension points</span></span></th>
    <th><span data-ttu-id="4af34-327">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="4af34-327">API requirement sets</span></span></th>
    <th><span data-ttu-id="4af34-328"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="4af34-328"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-329">Office Online</span><span class="sxs-lookup"><span data-stu-id="4af34-329">Office Online</span></span></td>
    <td> <span data-ttu-id="4af34-330">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="4af34-330">- Mail Read</span></span><br><span data-ttu-id="4af34-331">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="4af34-331">
      - Mail Compose</span></span><br><span data-ttu-id="4af34-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4af34-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4af34-333">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-333">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4af34-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4af34-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4af34-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4af34-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4af34-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4af34-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4af34-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4af34-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4af34-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4af34-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="4af34-339">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4af34-339">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="4af34-340">Não disponível</span><span class="sxs-lookup"><span data-stu-id="4af34-340">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-341">Office 365 para Windows</span><span class="sxs-lookup"><span data-stu-id="4af34-341">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="4af34-342">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="4af34-342">- Mail Read</span></span><br><span data-ttu-id="4af34-343">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="4af34-343">
      - Mail Compose</span></span><br><span data-ttu-id="4af34-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4af34-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="4af34-345">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="4af34-345">
      - Modules</span></span></td>
    <td> <span data-ttu-id="4af34-346">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-346">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4af34-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4af34-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4af34-348">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4af34-348">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4af34-349">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4af34-349">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4af34-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4af34-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4af34-351">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4af34-351">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="4af34-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4af34-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="4af34-353">Não disponível</span><span class="sxs-lookup"><span data-stu-id="4af34-353">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-354">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="4af34-354">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="4af34-355">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="4af34-355">- Mail Read</span></span><br><span data-ttu-id="4af34-356">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="4af34-356">
      - Mail Compose</span></span><br><span data-ttu-id="4af34-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4af34-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="4af34-358">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="4af34-358">
      - Modules</span></span></td>
    <td> <span data-ttu-id="4af34-359">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-359">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4af34-360">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4af34-360">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4af34-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4af34-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4af34-362">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4af34-362">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4af34-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4af34-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4af34-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4af34-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="4af34-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4af34-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="4af34-366">Não disponível</span><span class="sxs-lookup"><span data-stu-id="4af34-366">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-367">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="4af34-367">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="4af34-368">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="4af34-368">- Mail Read</span></span><br><span data-ttu-id="4af34-369">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="4af34-369">
      - Mail Compose</span></span><br><span data-ttu-id="4af34-370">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4af34-370">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="4af34-371">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="4af34-371">
      - Modules</span></span></td>
    <td> <span data-ttu-id="4af34-372">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-372">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4af34-373">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4af34-373">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4af34-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4af34-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4af34-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="4af34-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="4af34-376">Não disponível</span><span class="sxs-lookup"><span data-stu-id="4af34-376">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-377">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="4af34-377">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="4af34-378">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="4af34-378">- Mail Read</span></span><br><span data-ttu-id="4af34-379">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="4af34-379">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="4af34-380">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-380">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4af34-381">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4af34-381">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4af34-382">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="4af34-382">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4af34-383">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="4af34-383">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="4af34-384">Não disponível</span><span class="sxs-lookup"><span data-stu-id="4af34-384">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-385">Office 365 para iOS</span><span class="sxs-lookup"><span data-stu-id="4af34-385">See the Office 365 SDK for iOS.</span></span></td>
    <td> <span data-ttu-id="4af34-386">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="4af34-386">- Mail Read</span></span><br><span data-ttu-id="4af34-387">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4af34-387">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4af34-388">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-388">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4af34-389">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4af34-389">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4af34-390">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4af34-390">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4af34-391">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4af34-391">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4af34-392">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4af34-392">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="4af34-393">Não disponível</span><span class="sxs-lookup"><span data-stu-id="4af34-393">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-394">Office 365 para Mac</span><span class="sxs-lookup"><span data-stu-id="4af34-394">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="4af34-395">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="4af34-395">- Mail Read</span></span><br><span data-ttu-id="4af34-396">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="4af34-396">
      - Mail Compose</span></span><br><span data-ttu-id="4af34-397">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4af34-397">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4af34-398">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-398">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4af34-399">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4af34-399">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4af34-400">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4af34-400">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4af34-401">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4af34-401">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4af34-402">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4af34-402">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4af34-403">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4af34-403">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="4af34-404">Não disponível</span><span class="sxs-lookup"><span data-stu-id="4af34-404">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-405">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="4af34-405">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="4af34-406">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="4af34-406">- Mail Read</span></span><br><span data-ttu-id="4af34-407">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="4af34-407">
      - Mail Compose</span></span><br><span data-ttu-id="4af34-408">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4af34-408">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4af34-409">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-409">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4af34-410">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4af34-410">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4af34-411">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4af34-411">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4af34-412">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4af34-412">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4af34-413">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4af34-413">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4af34-414">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4af34-414">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="4af34-415">Não disponível</span><span class="sxs-lookup"><span data-stu-id="4af34-415">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-416">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="4af34-416">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="4af34-417">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="4af34-417">- Mail Read</span></span><br><span data-ttu-id="4af34-418">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="4af34-418">
      - Mail Compose</span></span><br><span data-ttu-id="4af34-419">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4af34-419">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4af34-420">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-420">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4af34-421">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4af34-421">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4af34-422">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4af34-422">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4af34-423">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4af34-423">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4af34-424">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4af34-424">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4af34-425">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4af34-425">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="4af34-426">Não disponível</span><span class="sxs-lookup"><span data-stu-id="4af34-426">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-427">Office 365 para Android</span><span class="sxs-lookup"><span data-stu-id="4af34-427">Office 365 SDK for Android</span></span></td>
    <td> <span data-ttu-id="4af34-428">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="4af34-428">- Mail Read</span></span><br><span data-ttu-id="4af34-429">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4af34-429">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4af34-430">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-430">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4af34-431">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4af34-431">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4af34-432">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4af34-432">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4af34-433">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4af34-433">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4af34-434">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4af34-434">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="4af34-435">Não disponível</span><span class="sxs-lookup"><span data-stu-id="4af34-435">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="4af34-436">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="4af34-436">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="4af34-437">Word</span><span class="sxs-lookup"><span data-stu-id="4af34-437">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4af34-438">Plataforma</span><span class="sxs-lookup"><span data-stu-id="4af34-438">Platform</span></span></th>
    <th><span data-ttu-id="4af34-439">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="4af34-439">Extension points</span></span></th>
    <th><span data-ttu-id="4af34-440">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="4af34-440">API requirement sets</span></span></th>
    <th><span data-ttu-id="4af34-441"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="4af34-441"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-442">Office Online</span><span class="sxs-lookup"><span data-stu-id="4af34-442">Office Online</span></span></td>
    <td> <span data-ttu-id="4af34-443">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="4af34-443">- TaskPane</span></span><br><span data-ttu-id="4af34-444">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4af34-444">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4af34-445">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-445">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4af34-446">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4af34-446">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4af34-447">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4af34-447">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4af34-448">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-448">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4af34-449">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-449">- BindingEvents</span></span><br><span data-ttu-id="4af34-450">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4af34-450">
         - CustomXmlParts</span></span><br><span data-ttu-id="4af34-451">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-451">
         - DocumentEvents</span></span><br><span data-ttu-id="4af34-452">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4af34-452">
         - File</span></span><br><span data-ttu-id="4af34-453">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-453">
         - HtmlCoercion</span></span><br><span data-ttu-id="4af34-454">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-454">
         - ImageCoercion</span></span><br><span data-ttu-id="4af34-455">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-455">
         - MatrixBindings</span></span><br><span data-ttu-id="4af34-456">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-456">
         - MatrixCoercion</span></span><br><span data-ttu-id="4af34-457">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-457">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4af34-458">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4af34-458">
         - PdfFile</span></span><br><span data-ttu-id="4af34-459">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="4af34-459">
         - Selection</span></span><br><span data-ttu-id="4af34-460">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="4af34-460">
         - Settings</span></span><br><span data-ttu-id="4af34-461">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-461">
         - TableBindings</span></span><br><span data-ttu-id="4af34-462">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-462">
         - TableCoercion</span></span><br><span data-ttu-id="4af34-463">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-463">
         - TextBindings</span></span><br><span data-ttu-id="4af34-464">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-464">
         - TextCoercion</span></span><br><span data-ttu-id="4af34-465">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4af34-465">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-466">Office 365 para Windows</span><span class="sxs-lookup"><span data-stu-id="4af34-466">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="4af34-467">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="4af34-467">- TaskPane</span></span><br><span data-ttu-id="4af34-468">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4af34-468">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4af34-469">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-469">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4af34-470">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4af34-470">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4af34-471">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4af34-471">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4af34-472">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-472">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4af34-473">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-473">- BindingEvents</span></span><br><span data-ttu-id="4af34-474">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4af34-474">
         - CompressedFile</span></span><br><span data-ttu-id="4af34-475">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4af34-475">
         - CustomXmlParts</span></span><br><span data-ttu-id="4af34-476">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-476">
         - DocumentEvents</span></span><br><span data-ttu-id="4af34-477">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4af34-477">
         - File</span></span><br><span data-ttu-id="4af34-478">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-478">
         - HtmlCoercion</span></span><br><span data-ttu-id="4af34-479">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-479">
         - ImageCoercion</span></span><br><span data-ttu-id="4af34-480">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-480">
         - MatrixBindings</span></span><br><span data-ttu-id="4af34-481">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-481">
         - MatrixCoercion</span></span><br><span data-ttu-id="4af34-482">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-482">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4af34-483">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4af34-483">
         - PdfFile</span></span><br><span data-ttu-id="4af34-484">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="4af34-484">
         - Selection</span></span><br><span data-ttu-id="4af34-485">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="4af34-485">
         - Settings</span></span><br><span data-ttu-id="4af34-486">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-486">
         - TableBindings</span></span><br><span data-ttu-id="4af34-487">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-487">
         - TableCoercion</span></span><br><span data-ttu-id="4af34-488">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-488">
         - TextBindings</span></span><br><span data-ttu-id="4af34-489">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-489">
         - TextCoercion</span></span><br><span data-ttu-id="4af34-490">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4af34-490">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-491">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="4af34-491">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="4af34-492">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="4af34-492">- TaskPane</span></span><br><span data-ttu-id="4af34-493">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4af34-493">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4af34-494">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-494">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4af34-495">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4af34-495">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4af34-496">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4af34-496">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4af34-497">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-497">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4af34-498">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-498">- BindingEvents</span></span><br><span data-ttu-id="4af34-499">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4af34-499">
         - CompressedFile</span></span><br><span data-ttu-id="4af34-500">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4af34-500">
         - CustomXmlParts</span></span><br><span data-ttu-id="4af34-501">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-501">
         - DocumentEvents</span></span><br><span data-ttu-id="4af34-502">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4af34-502">
         - File</span></span><br><span data-ttu-id="4af34-503">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-503">
         - HtmlCoercion</span></span><br><span data-ttu-id="4af34-504">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-504">
         - ImageCoercion</span></span><br><span data-ttu-id="4af34-505">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-505">
         - MatrixBindings</span></span><br><span data-ttu-id="4af34-506">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-506">
         - MatrixCoercion</span></span><br><span data-ttu-id="4af34-507">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-507">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4af34-508">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4af34-508">
         - PdfFile</span></span><br><span data-ttu-id="4af34-509">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="4af34-509">
         - Selection</span></span><br><span data-ttu-id="4af34-510">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="4af34-510">
         - Settings</span></span><br><span data-ttu-id="4af34-511">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-511">
         - TableBindings</span></span><br><span data-ttu-id="4af34-512">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-512">
         - TableCoercion</span></span><br><span data-ttu-id="4af34-513">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-513">
         - TextBindings</span></span><br><span data-ttu-id="4af34-514">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-514">
         - TextCoercion</span></span><br><span data-ttu-id="4af34-515">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4af34-515">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-516">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="4af34-516">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="4af34-517">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="4af34-517">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4af34-518">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-518">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4af34-519">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="4af34-519">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="4af34-520">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-520">- BindingEvents</span></span><br><span data-ttu-id="4af34-521">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4af34-521">
         - CompressedFile</span></span><br><span data-ttu-id="4af34-522">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4af34-522">
         - CustomXmlParts</span></span><br><span data-ttu-id="4af34-523">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-523">
         - DocumentEvents</span></span><br><span data-ttu-id="4af34-524">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4af34-524">
         - File</span></span><br><span data-ttu-id="4af34-525">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-525">
         - HtmlCoercion</span></span><br><span data-ttu-id="4af34-526">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-526">
         - ImageCoercion</span></span><br><span data-ttu-id="4af34-527">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-527">
         - MatrixBindings</span></span><br><span data-ttu-id="4af34-528">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-528">
         - MatrixCoercion</span></span><br><span data-ttu-id="4af34-529">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-529">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4af34-530">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4af34-530">
         - PdfFile</span></span><br><span data-ttu-id="4af34-531">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="4af34-531">
         - Selection</span></span><br><span data-ttu-id="4af34-532">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="4af34-532">
         - Settings</span></span><br><span data-ttu-id="4af34-533">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-533">
         - TableBindings</span></span><br><span data-ttu-id="4af34-534">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-534">
         - TableCoercion</span></span><br><span data-ttu-id="4af34-535">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-535">
         - TextBindings</span></span><br><span data-ttu-id="4af34-536">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-536">
         - TextCoercion</span></span><br><span data-ttu-id="4af34-537">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4af34-537">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-538">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="4af34-538">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="4af34-539">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="4af34-539">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4af34-540">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="4af34-540">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="4af34-541">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-541">- BindingEvents</span></span><br><span data-ttu-id="4af34-542">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4af34-542">
         - CompressedFile</span></span><br><span data-ttu-id="4af34-543">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4af34-543">
         - CustomXmlParts</span></span><br><span data-ttu-id="4af34-544">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-544">
         - DocumentEvents</span></span><br><span data-ttu-id="4af34-545">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4af34-545">
         - File</span></span><br><span data-ttu-id="4af34-546">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-546">
         - HtmlCoercion</span></span><br><span data-ttu-id="4af34-547">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-547">
         - ImageCoercion</span></span><br><span data-ttu-id="4af34-548">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-548">
         - MatrixBindings</span></span><br><span data-ttu-id="4af34-549">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-549">
         - MatrixCoercion</span></span><br><span data-ttu-id="4af34-550">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-550">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4af34-551">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4af34-551">
         - PdfFile</span></span><br><span data-ttu-id="4af34-552">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="4af34-552">
         - Selection</span></span><br><span data-ttu-id="4af34-553">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="4af34-553">
         - Settings</span></span><br><span data-ttu-id="4af34-554">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-554">
         - TableBindings</span></span><br><span data-ttu-id="4af34-555">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-555">
         - TableCoercion</span></span><br><span data-ttu-id="4af34-556">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-556">
         - TextBindings</span></span><br><span data-ttu-id="4af34-557">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-557">
         - TextCoercion</span></span><br><span data-ttu-id="4af34-558">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4af34-558">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-559">Office 365 para iPad</span><span class="sxs-lookup"><span data-stu-id="4af34-559">Office 365 for iPad</span></span></td>
    <td> <span data-ttu-id="4af34-560">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="4af34-560">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4af34-561">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-561">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4af34-562">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4af34-562">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4af34-563">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4af34-563">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4af34-564">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="4af34-564">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="4af34-565">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-565">- BindingEvents</span></span><br><span data-ttu-id="4af34-566">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4af34-566">
         - CompressedFile</span></span><br><span data-ttu-id="4af34-567">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4af34-567">
         - CustomXmlParts</span></span><br><span data-ttu-id="4af34-568">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-568">
         - DocumentEvents</span></span><br><span data-ttu-id="4af34-569">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4af34-569">
         - File</span></span><br><span data-ttu-id="4af34-570">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-570">
         - HtmlCoercion</span></span><br><span data-ttu-id="4af34-571">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-571">
         - ImageCoercion</span></span><br><span data-ttu-id="4af34-572">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-572">
         - MatrixBindings</span></span><br><span data-ttu-id="4af34-573">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-573">
         - MatrixCoercion</span></span><br><span data-ttu-id="4af34-574">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-574">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4af34-575">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4af34-575">
         - PdfFile</span></span><br><span data-ttu-id="4af34-576">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="4af34-576">
         - Selection</span></span><br><span data-ttu-id="4af34-577">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="4af34-577">
         - Settings</span></span><br><span data-ttu-id="4af34-578">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-578">
         - TableBindings</span></span><br><span data-ttu-id="4af34-579">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-579">
         - TableCoercion</span></span><br><span data-ttu-id="4af34-580">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-580">
         - TextBindings</span></span><br><span data-ttu-id="4af34-581">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-581">
         - TextCoercion</span></span><br><span data-ttu-id="4af34-582">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4af34-582">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-583">Office 365 para Mac</span><span class="sxs-lookup"><span data-stu-id="4af34-583">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="4af34-584">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="4af34-584">- TaskPane</span></span><br><span data-ttu-id="4af34-585">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4af34-585">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4af34-586">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-586">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4af34-587">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4af34-587">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4af34-588">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4af34-588">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4af34-589">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="4af34-589">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="4af34-590">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-590">- BindingEvents</span></span><br><span data-ttu-id="4af34-591">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4af34-591">
         - CompressedFile</span></span><br><span data-ttu-id="4af34-592">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4af34-592">
         - CustomXmlParts</span></span><br><span data-ttu-id="4af34-593">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-593">
         - DocumentEvents</span></span><br><span data-ttu-id="4af34-594">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4af34-594">
         - File</span></span><br><span data-ttu-id="4af34-595">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-595">
         - HtmlCoercion</span></span><br><span data-ttu-id="4af34-596">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-596">
         - ImageCoercion</span></span><br><span data-ttu-id="4af34-597">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-597">
         - MatrixBindings</span></span><br><span data-ttu-id="4af34-598">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-598">
         - MatrixCoercion</span></span><br><span data-ttu-id="4af34-599">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-599">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4af34-600">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4af34-600">
         - PdfFile</span></span><br><span data-ttu-id="4af34-601">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="4af34-601">
         - Selection</span></span><br><span data-ttu-id="4af34-602">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="4af34-602">
         - Settings</span></span><br><span data-ttu-id="4af34-603">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-603">
         - TableBindings</span></span><br><span data-ttu-id="4af34-604">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-604">
         - TableCoercion</span></span><br><span data-ttu-id="4af34-605">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-605">
         - TextBindings</span></span><br><span data-ttu-id="4af34-606">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-606">
         - TextCoercion</span></span><br><span data-ttu-id="4af34-607">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4af34-607">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-608">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="4af34-608">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="4af34-609">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="4af34-609">- TaskPane</span></span><br><span data-ttu-id="4af34-610">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4af34-610">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4af34-611">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-611">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4af34-612">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4af34-612">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4af34-613">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4af34-613">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4af34-614">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="4af34-614">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="4af34-615">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-615">- BindingEvents</span></span><br><span data-ttu-id="4af34-616">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4af34-616">
         - CompressedFile</span></span><br><span data-ttu-id="4af34-617">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4af34-617">
         - CustomXmlParts</span></span><br><span data-ttu-id="4af34-618">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-618">
         - DocumentEvents</span></span><br><span data-ttu-id="4af34-619">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4af34-619">
         - File</span></span><br><span data-ttu-id="4af34-620">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-620">
         - HtmlCoercion</span></span><br><span data-ttu-id="4af34-621">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-621">
         - ImageCoercion</span></span><br><span data-ttu-id="4af34-622">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-622">
         - MatrixBindings</span></span><br><span data-ttu-id="4af34-623">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-623">
         - MatrixCoercion</span></span><br><span data-ttu-id="4af34-624">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-624">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4af34-625">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4af34-625">
         - PdfFile</span></span><br><span data-ttu-id="4af34-626">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="4af34-626">
         - Selection</span></span><br><span data-ttu-id="4af34-627">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="4af34-627">
         - Settings</span></span><br><span data-ttu-id="4af34-628">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-628">
         - TableBindings</span></span><br><span data-ttu-id="4af34-629">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-629">
         - TableCoercion</span></span><br><span data-ttu-id="4af34-630">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-630">
         - TextBindings</span></span><br><span data-ttu-id="4af34-631">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-631">
         - TextCoercion</span></span><br><span data-ttu-id="4af34-632">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4af34-632">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-633">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="4af34-633">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="4af34-634">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="4af34-634">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4af34-635">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-635">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4af34-636">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="4af34-636">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="4af34-637">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-637">- BindingEvents</span></span><br><span data-ttu-id="4af34-638">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4af34-638">
         - CompressedFile</span></span><br><span data-ttu-id="4af34-639">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4af34-639">
         - CustomXmlParts</span></span><br><span data-ttu-id="4af34-640">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-640">
         - DocumentEvents</span></span><br><span data-ttu-id="4af34-641">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4af34-641">
         - File</span></span><br><span data-ttu-id="4af34-642">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-642">
         - HtmlCoercion</span></span><br><span data-ttu-id="4af34-643">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-643">
         - ImageCoercion</span></span><br><span data-ttu-id="4af34-644">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-644">
         - MatrixBindings</span></span><br><span data-ttu-id="4af34-645">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-645">
         - MatrixCoercion</span></span><br><span data-ttu-id="4af34-646">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-646">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4af34-647">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4af34-647">
         - PdfFile</span></span><br><span data-ttu-id="4af34-648">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="4af34-648">
         - Selection</span></span><br><span data-ttu-id="4af34-649">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="4af34-649">
         - Settings</span></span><br><span data-ttu-id="4af34-650">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-650">
         - TableBindings</span></span><br><span data-ttu-id="4af34-651">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-651">
         - TableCoercion</span></span><br><span data-ttu-id="4af34-652">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4af34-652">
         - TextBindings</span></span><br><span data-ttu-id="4af34-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-653">
         - TextCoercion</span></span><br><span data-ttu-id="4af34-654">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4af34-654">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="4af34-655">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="4af34-655">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="4af34-656">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="4af34-656">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4af34-657">Plataforma</span><span class="sxs-lookup"><span data-stu-id="4af34-657">Platform</span></span></th>
    <th><span data-ttu-id="4af34-658">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="4af34-658">Extension points</span></span></th>
    <th><span data-ttu-id="4af34-659">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="4af34-659">API requirement sets</span></span></th>
    <th><span data-ttu-id="4af34-660"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="4af34-660"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-661">Office Online</span><span class="sxs-lookup"><span data-stu-id="4af34-661">Office Online</span></span></td>
    <td> <span data-ttu-id="4af34-662">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="4af34-662">- Content</span></span><br><span data-ttu-id="4af34-663">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="4af34-663">
         - TaskPane</span></span><br><span data-ttu-id="4af34-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4af34-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4af34-665">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-665">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4af34-666">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4af34-666">- ActiveView</span></span><br><span data-ttu-id="4af34-667">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4af34-667">
         - CompressedFile</span></span><br><span data-ttu-id="4af34-668">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-668">
         - DocumentEvents</span></span><br><span data-ttu-id="4af34-669">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4af34-669">
         - File</span></span><br><span data-ttu-id="4af34-670">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-670">
         - ImageCoercion</span></span><br><span data-ttu-id="4af34-671">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4af34-671">
         - PdfFile</span></span><br><span data-ttu-id="4af34-672">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="4af34-672">
         - Selection</span></span><br><span data-ttu-id="4af34-673">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="4af34-673">
         - Settings</span></span><br><span data-ttu-id="4af34-674">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-674">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-675">Office 365 para Windows</span><span class="sxs-lookup"><span data-stu-id="4af34-675">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="4af34-676">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="4af34-676">- Content</span></span><br><span data-ttu-id="4af34-677">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="4af34-677">
         - TaskPane</span></span><br><span data-ttu-id="4af34-678">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4af34-678">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4af34-679">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-679">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4af34-680">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4af34-680">- ActiveView</span></span><br><span data-ttu-id="4af34-681">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4af34-681">
         - CompressedFile</span></span><br><span data-ttu-id="4af34-682">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-682">
         - DocumentEvents</span></span><br><span data-ttu-id="4af34-683">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4af34-683">
         - File</span></span><br><span data-ttu-id="4af34-684">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-684">
         - ImageCoercion</span></span><br><span data-ttu-id="4af34-685">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4af34-685">
         - PdfFile</span></span><br><span data-ttu-id="4af34-686">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="4af34-686">
         - Selection</span></span><br><span data-ttu-id="4af34-687">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="4af34-687">
         - Settings</span></span><br><span data-ttu-id="4af34-688">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-688">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-689">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="4af34-689">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="4af34-690">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="4af34-690">- Content</span></span><br><span data-ttu-id="4af34-691">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="4af34-691">
         - TaskPane</span></span><br><span data-ttu-id="4af34-692">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4af34-692">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4af34-693">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-693">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4af34-694">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4af34-694">- ActiveView</span></span><br><span data-ttu-id="4af34-695">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4af34-695">
         - CompressedFile</span></span><br><span data-ttu-id="4af34-696">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-696">
         - DocumentEvents</span></span><br><span data-ttu-id="4af34-697">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4af34-697">
         - File</span></span><br><span data-ttu-id="4af34-698">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-698">
         - ImageCoercion</span></span><br><span data-ttu-id="4af34-699">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4af34-699">
         - PdfFile</span></span><br><span data-ttu-id="4af34-700">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="4af34-700">
         - Selection</span></span><br><span data-ttu-id="4af34-701">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="4af34-701">
         - Settings</span></span><br><span data-ttu-id="4af34-702">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-702">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-703">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="4af34-703">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="4af34-704">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="4af34-704">- Content</span></span><br><span data-ttu-id="4af34-705">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="4af34-705">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="4af34-706">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="4af34-706">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="4af34-707">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4af34-707">- ActiveView</span></span><br><span data-ttu-id="4af34-708">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4af34-708">
         - CompressedFile</span></span><br><span data-ttu-id="4af34-709">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-709">
         - DocumentEvents</span></span><br><span data-ttu-id="4af34-710">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4af34-710">
         - File</span></span><br><span data-ttu-id="4af34-711">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-711">
         - ImageCoercion</span></span><br><span data-ttu-id="4af34-712">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4af34-712">
         - PdfFile</span></span><br><span data-ttu-id="4af34-713">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="4af34-713">
         - Selection</span></span><br><span data-ttu-id="4af34-714">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="4af34-714">
         - Settings</span></span><br><span data-ttu-id="4af34-715">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-715">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-716">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="4af34-716">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="4af34-717">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="4af34-717">- Content</span></span><br><span data-ttu-id="4af34-718">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="4af34-718">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="4af34-719">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="4af34-719">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="4af34-720">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4af34-720">- ActiveView</span></span><br><span data-ttu-id="4af34-721">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4af34-721">
         - CompressedFile</span></span><br><span data-ttu-id="4af34-722">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-722">
         - DocumentEvents</span></span><br><span data-ttu-id="4af34-723">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4af34-723">
         - File</span></span><br><span data-ttu-id="4af34-724">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-724">
         - ImageCoercion</span></span><br><span data-ttu-id="4af34-725">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4af34-725">
         - PdfFile</span></span><br><span data-ttu-id="4af34-726">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="4af34-726">
         - Selection</span></span><br><span data-ttu-id="4af34-727">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="4af34-727">
         - Settings</span></span><br><span data-ttu-id="4af34-728">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-728">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-729">Office 365 para iPad</span><span class="sxs-lookup"><span data-stu-id="4af34-729">Office 365 for iPad</span></span></td>
    <td> <span data-ttu-id="4af34-730">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="4af34-730">- Content</span></span><br><span data-ttu-id="4af34-731">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="4af34-731">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="4af34-732">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-732">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="4af34-733">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4af34-733">- ActiveView</span></span><br><span data-ttu-id="4af34-734">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4af34-734">
         - CompressedFile</span></span><br><span data-ttu-id="4af34-735">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-735">
         - DocumentEvents</span></span><br><span data-ttu-id="4af34-736">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4af34-736">
         - File</span></span><br><span data-ttu-id="4af34-737">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4af34-737">
         - PdfFile</span></span><br><span data-ttu-id="4af34-738">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="4af34-738">
         - Selection</span></span><br><span data-ttu-id="4af34-739">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="4af34-739">
         - Settings</span></span><br><span data-ttu-id="4af34-740">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-740">
         - TextCoercion</span></span><br><span data-ttu-id="4af34-741">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-741">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-742">Office 365 para Mac</span><span class="sxs-lookup"><span data-stu-id="4af34-742">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="4af34-743">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="4af34-743">- Content</span></span><br><span data-ttu-id="4af34-744">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="4af34-744">
         - TaskPane</span></span><br><span data-ttu-id="4af34-745">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4af34-745">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4af34-746">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-746">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4af34-747">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4af34-747">- ActiveView</span></span><br><span data-ttu-id="4af34-748">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4af34-748">
         - CompressedFile</span></span><br><span data-ttu-id="4af34-749">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-749">
         - DocumentEvents</span></span><br><span data-ttu-id="4af34-750">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4af34-750">
         - File</span></span><br><span data-ttu-id="4af34-751">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-751">
         - ImageCoercion</span></span><br><span data-ttu-id="4af34-752">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4af34-752">
         - PdfFile</span></span><br><span data-ttu-id="4af34-753">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="4af34-753">
         - Selection</span></span><br><span data-ttu-id="4af34-754">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="4af34-754">
         - Settings</span></span><br><span data-ttu-id="4af34-755">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-755">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-756">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="4af34-756">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="4af34-757">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="4af34-757">- Content</span></span><br><span data-ttu-id="4af34-758">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="4af34-758">
         - TaskPane</span></span><br><span data-ttu-id="4af34-759">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4af34-759">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4af34-760">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-760">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4af34-761">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4af34-761">- ActiveView</span></span><br><span data-ttu-id="4af34-762">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4af34-762">
         - CompressedFile</span></span><br><span data-ttu-id="4af34-763">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-763">
         - DocumentEvents</span></span><br><span data-ttu-id="4af34-764">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4af34-764">
         - File</span></span><br><span data-ttu-id="4af34-765">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-765">
         - ImageCoercion</span></span><br><span data-ttu-id="4af34-766">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4af34-766">
         - PdfFile</span></span><br><span data-ttu-id="4af34-767">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="4af34-767">
         - Selection</span></span><br><span data-ttu-id="4af34-768">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="4af34-768">
         - Settings</span></span><br><span data-ttu-id="4af34-769">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-769">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-770">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="4af34-770">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="4af34-771">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="4af34-771">- Content</span></span><br><span data-ttu-id="4af34-772">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="4af34-772">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="4af34-773">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="4af34-773">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="4af34-774">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4af34-774">- ActiveView</span></span><br><span data-ttu-id="4af34-775">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4af34-775">
         - CompressedFile</span></span><br><span data-ttu-id="4af34-776">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-776">
         - DocumentEvents</span></span><br><span data-ttu-id="4af34-777">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4af34-777">
         - File</span></span><br><span data-ttu-id="4af34-778">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-778">
         - ImageCoercion</span></span><br><span data-ttu-id="4af34-779">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4af34-779">
         - PdfFile</span></span><br><span data-ttu-id="4af34-780">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="4af34-780">
         - Selection</span></span><br><span data-ttu-id="4af34-781">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="4af34-781">
         - Settings</span></span><br><span data-ttu-id="4af34-782">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-782">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="4af34-783">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="4af34-783">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="4af34-784">OneNote</span><span class="sxs-lookup"><span data-stu-id="4af34-784">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4af34-785">Plataforma</span><span class="sxs-lookup"><span data-stu-id="4af34-785">Platform</span></span></th>
    <th><span data-ttu-id="4af34-786">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="4af34-786">Extension points</span></span></th>
    <th><span data-ttu-id="4af34-787">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="4af34-787">API requirement sets</span></span></th>
    <th><span data-ttu-id="4af34-788"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="4af34-788"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-789">Office Online</span><span class="sxs-lookup"><span data-stu-id="4af34-789">Office Online</span></span></td>
    <td> <span data-ttu-id="4af34-790">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="4af34-790">- Content</span></span><br><span data-ttu-id="4af34-791">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="4af34-791">
         - TaskPane</span></span><br><span data-ttu-id="4af34-792">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4af34-792">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4af34-793">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-793">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="4af34-794">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-794">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4af34-795">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4af34-795">- DocumentEvents</span></span><br><span data-ttu-id="4af34-796">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-796">
         - HtmlCoercion</span></span><br><span data-ttu-id="4af34-797">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-797">
         - ImageCoercion</span></span><br><span data-ttu-id="4af34-798">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="4af34-798">
         - Settings</span></span><br><span data-ttu-id="4af34-799">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-799">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="4af34-800">Project</span><span class="sxs-lookup"><span data-stu-id="4af34-800">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4af34-801">Plataforma</span><span class="sxs-lookup"><span data-stu-id="4af34-801">Platform</span></span></th>
    <th><span data-ttu-id="4af34-802">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="4af34-802">Extension points</span></span></th>
    <th><span data-ttu-id="4af34-803">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="4af34-803">API requirement sets</span></span></th>
    <th><span data-ttu-id="4af34-804"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="4af34-804"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-805">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="4af34-805">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="4af34-806">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="4af34-806">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4af34-807">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-807">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4af34-808">- Seleção</span><span class="sxs-lookup"><span data-stu-id="4af34-808">- Selection</span></span><br><span data-ttu-id="4af34-809">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-809">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-810">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="4af34-810">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="4af34-811">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="4af34-811">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4af34-812">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-812">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4af34-813">- Seleção</span><span class="sxs-lookup"><span data-stu-id="4af34-813">- Selection</span></span><br><span data-ttu-id="4af34-814">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-814">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4af34-815">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="4af34-815">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="4af34-816">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="4af34-816">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4af34-817">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4af34-817">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4af34-818">- Seleção</span><span class="sxs-lookup"><span data-stu-id="4af34-818">- Selection</span></span><br><span data-ttu-id="4af34-819">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4af34-819">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="4af34-820">Confira também</span><span class="sxs-lookup"><span data-stu-id="4af34-820">See also</span></span>

- [<span data-ttu-id="4af34-821">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="4af34-821">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="4af34-822">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="4af34-822">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="4af34-823">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="4af34-823">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="4af34-824">Referência da API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="4af34-824">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
