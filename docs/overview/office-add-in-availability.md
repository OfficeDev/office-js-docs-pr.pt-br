---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos compatíveis com Excel, Word, Outlook, PowerPoint, OneNote e Project.
ms.date: 11/07/2018
localization_priority: Priority
ms.openlocfilehash: 9f8b94483d22f24dcb0a6a2ad99df6167533133f
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388336"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="61e6f-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="61e6f-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="61e6f-104">Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="61e6f-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="61e6f-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que são compatíveis atualmente com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="61e6f-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="61e6f-p102">O número do build do Office 2016 instalado via MSI é 16.0.4266.1001. Esta versão só contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="61e6f-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="61e6f-108">Excel</span><span class="sxs-lookup"><span data-stu-id="61e6f-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="61e6f-109">Plataforma</span><span class="sxs-lookup"><span data-stu-id="61e6f-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="61e6f-110">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="61e6f-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="61e6f-111">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="61e6f-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="61e6f-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="61e6f-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="61e6f-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="61e6f-113">Office Online</span></span></td>
    <td> <span data-ttu-id="61e6f-114">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="61e6f-114">- TaskPane</span></span><br><span data-ttu-id="61e6f-115">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="61e6f-115">
        - Content</span></span><br><span data-ttu-id="61e6f-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="61e6f-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="61e6f-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="61e6f-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="61e6f-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="61e6f-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="61e6f-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="61e6f-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="61e6f-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="61e6f-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="61e6f-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="61e6f-126">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="61e6f-126">
        - BindingEvents</span></span><br><span data-ttu-id="61e6f-127">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-127">
        - CompressedFile</span></span><br><span data-ttu-id="61e6f-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="61e6f-128">
        - DocumentEvents</span></span><br><span data-ttu-id="61e6f-129">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="61e6f-129">
        - File</span></span><br><span data-ttu-id="61e6f-130">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-130">
        - MatrixBindings</span></span><br><span data-ttu-id="61e6f-131">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-131">
        - MatrixCoercion</span></span><br><span data-ttu-id="61e6f-132">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="61e6f-132">
        - Selection</span></span><br><span data-ttu-id="61e6f-133">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="61e6f-133">
        - Settings</span></span><br><span data-ttu-id="61e6f-134">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-134">
        - TableBindings</span></span><br><span data-ttu-id="61e6f-135">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-135">
        - TableCoercion</span></span><br><span data-ttu-id="61e6f-136">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-136">
        - TextBindings</span></span><br><span data-ttu-id="61e6f-137">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-137">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="61e6f-138">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="61e6f-138">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="61e6f-139">
        - TaskPane</span><span class="sxs-lookup"><span data-stu-id="61e6f-139">
        - TaskPane</span></span><br><span data-ttu-id="61e6f-140">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="61e6f-140">
        - Content</span></span></td>
    <td>  <span data-ttu-id="61e6f-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="61e6f-142">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="61e6f-142">
        - BindingEvents</span></span><br><span data-ttu-id="61e6f-143">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-143">
        - CompressedFile</span></span><br><span data-ttu-id="61e6f-144">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="61e6f-144">
        - DocumentEvents</span></span><br><span data-ttu-id="61e6f-145">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="61e6f-145">
        - File</span></span><br><span data-ttu-id="61e6f-146">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-146">
        - ImageCoercion</span></span><br><span data-ttu-id="61e6f-147">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-147">
        - MatrixBindings</span></span><br><span data-ttu-id="61e6f-148">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-148">
        - MatrixCoercion</span></span><br><span data-ttu-id="61e6f-149">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="61e6f-149">
        - Selection</span></span><br><span data-ttu-id="61e6f-150">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="61e6f-150">
        - Settings</span></span><br><span data-ttu-id="61e6f-151">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-151">
        - TableBindings</span></span><br><span data-ttu-id="61e6f-152">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-152">
        - TableCoercion</span></span><br><span data-ttu-id="61e6f-153">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-153">
        - TextBindings</span></span><br><span data-ttu-id="61e6f-154">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-154">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="61e6f-155">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="61e6f-155">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="61e6f-156">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="61e6f-156">- TaskPane</span></span><br><span data-ttu-id="61e6f-157">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="61e6f-157">
        - Content</span></span><br><span data-ttu-id="61e6f-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="61e6f-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="61e6f-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="61e6f-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="61e6f-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="61e6f-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="61e6f-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="61e6f-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="61e6f-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="61e6f-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="61e6f-168">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="61e6f-168">- BindingEvents</span></span><br><span data-ttu-id="61e6f-169">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-169">
        - CompressedFile</span></span><br><span data-ttu-id="61e6f-170">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="61e6f-170">
        - DocumentEvents</span></span><br><span data-ttu-id="61e6f-171">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="61e6f-171">
        - File</span></span><br><span data-ttu-id="61e6f-172">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-172">
        - ImageCoercion</span></span><br><span data-ttu-id="61e6f-173">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-173">
        - MatrixBindings</span></span><br><span data-ttu-id="61e6f-174">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-174">
        - MatrixCoercion</span></span><br><span data-ttu-id="61e6f-175">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="61e6f-175">
        - Selection</span></span><br><span data-ttu-id="61e6f-176">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="61e6f-176">
        - Settings</span></span><br><span data-ttu-id="61e6f-177">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-177">
        - TableBindings</span></span><br><span data-ttu-id="61e6f-178">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-178">
        - TableCoercion</span></span><br><span data-ttu-id="61e6f-179">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-179">
        - TextBindings</span></span><br><span data-ttu-id="61e6f-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-180">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="61e6f-181">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="61e6f-181">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="61e6f-182">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="61e6f-182">- TaskPane</span></span><br><span data-ttu-id="61e6f-183">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="61e6f-183">
        - Content</span></span><br><span data-ttu-id="61e6f-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="61e6f-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="61e6f-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="61e6f-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="61e6f-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="61e6f-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="61e6f-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="61e6f-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="61e6f-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="61e6f-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="61e6f-194">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="61e6f-194">- BindingEvents</span></span><br><span data-ttu-id="61e6f-195">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-195">
        - CompressedFile</span></span><br><span data-ttu-id="61e6f-196">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="61e6f-196">
        - DocumentEvents</span></span><br><span data-ttu-id="61e6f-197">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="61e6f-197">
        - File</span></span><br><span data-ttu-id="61e6f-198">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-198">
        - ImageCoercion</span></span><br><span data-ttu-id="61e6f-199">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-199">
        - MatrixBindings</span></span><br><span data-ttu-id="61e6f-200">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-200">
        - MatrixCoercion</span></span><br><span data-ttu-id="61e6f-201">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="61e6f-201">
        - Selection</span></span><br><span data-ttu-id="61e6f-202">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="61e6f-202">
        - Settings</span></span><br><span data-ttu-id="61e6f-203">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-203">
        - TableBindings</span></span><br><span data-ttu-id="61e6f-204">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-204">
        - TableCoercion</span></span><br><span data-ttu-id="61e6f-205">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-205">
        - TextBindings</span></span><br><span data-ttu-id="61e6f-206">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-206">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="61e6f-207">Office para iPad</span><span class="sxs-lookup"><span data-stu-id="61e6f-207">Office for iPad</span></span></td>
    <td><span data-ttu-id="61e6f-208">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="61e6f-208">- TaskPane</span></span><br><span data-ttu-id="61e6f-209">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="61e6f-209">
        - Content</span></span></td>
    <td><span data-ttu-id="61e6f-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="61e6f-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="61e6f-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="61e6f-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="61e6f-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="61e6f-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="61e6f-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="61e6f-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="61e6f-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="61e6f-219">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="61e6f-219">- BindingEvents</span></span><br><span data-ttu-id="61e6f-220">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-220">
        - CompressedFile</span></span><br><span data-ttu-id="61e6f-221">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="61e6f-221">
        - DocumentEvents</span></span><br><span data-ttu-id="61e6f-222">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="61e6f-222">
        - File</span></span><br><span data-ttu-id="61e6f-223">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-223">
        - ImageCoercion</span></span><br><span data-ttu-id="61e6f-224">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-224">
        - MatrixBindings</span></span><br><span data-ttu-id="61e6f-225">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-225">
        - MatrixCoercion</span></span><br><span data-ttu-id="61e6f-226">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="61e6f-226">
        - Selection</span></span><br><span data-ttu-id="61e6f-227">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="61e6f-227">
        - Settings</span></span><br><span data-ttu-id="61e6f-228">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-228">
        - TableBindings</span></span><br><span data-ttu-id="61e6f-229">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-229">
        - TableCoercion</span></span><br><span data-ttu-id="61e6f-230">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-230">
        - TextBindings</span></span><br><span data-ttu-id="61e6f-231">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-231">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="61e6f-232">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="61e6f-232">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="61e6f-233">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="61e6f-233">- TaskPane</span></span><br><span data-ttu-id="61e6f-234">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="61e6f-234">
        - Content</span></span><br><span data-ttu-id="61e6f-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="61e6f-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="61e6f-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="61e6f-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="61e6f-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="61e6f-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="61e6f-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="61e6f-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="61e6f-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="61e6f-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="61e6f-245">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="61e6f-245">- BindingEvents</span></span><br><span data-ttu-id="61e6f-246">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-246">
        - CompressedFile</span></span><br><span data-ttu-id="61e6f-247">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="61e6f-247">
        - DocumentEvents</span></span><br><span data-ttu-id="61e6f-248">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="61e6f-248">
        - File</span></span><br><span data-ttu-id="61e6f-249">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-249">
        - ImageCoercion</span></span><br><span data-ttu-id="61e6f-250">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-250">
        - MatrixBindings</span></span><br><span data-ttu-id="61e6f-251">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-251">
        - MatrixCoercion</span></span><br><span data-ttu-id="61e6f-252">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-252">
        - PdfFile</span></span><br><span data-ttu-id="61e6f-253">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="61e6f-253">
        - Selection</span></span><br><span data-ttu-id="61e6f-254">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="61e6f-254">
        - Settings</span></span><br><span data-ttu-id="61e6f-255">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-255">
        - TableBindings</span></span><br><span data-ttu-id="61e6f-256">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-256">
        - TableCoercion</span></span><br><span data-ttu-id="61e6f-257">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-257">
        - TextBindings</span></span><br><span data-ttu-id="61e6f-258">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-258">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="61e6f-259">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="61e6f-259">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="61e6f-260">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="61e6f-260">- TaskPane</span></span><br><span data-ttu-id="61e6f-261">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="61e6f-261">
        - Content</span></span><br><span data-ttu-id="61e6f-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="61e6f-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="61e6f-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="61e6f-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="61e6f-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="61e6f-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="61e6f-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="61e6f-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="61e6f-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="61e6f-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="61e6f-272">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="61e6f-272">- BindingEvents</span></span><br><span data-ttu-id="61e6f-273">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-273">
        - CompressedFile</span></span><br><span data-ttu-id="61e6f-274">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="61e6f-274">
        - DocumentEvents</span></span><br><span data-ttu-id="61e6f-275">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="61e6f-275">
        - File</span></span><br><span data-ttu-id="61e6f-276">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-276">
        - ImageCoercion</span></span><br><span data-ttu-id="61e6f-277">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-277">
        - MatrixBindings</span></span><br><span data-ttu-id="61e6f-278">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-278">
        - MatrixCoercion</span></span><br><span data-ttu-id="61e6f-279">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-279">
        - PdfFile</span></span><br><span data-ttu-id="61e6f-280">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="61e6f-280">
        - Selection</span></span><br><span data-ttu-id="61e6f-281">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="61e6f-281">
        - Settings</span></span><br><span data-ttu-id="61e6f-282">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-282">
        - TableBindings</span></span><br><span data-ttu-id="61e6f-283">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-283">
        - TableCoercion</span></span><br><span data-ttu-id="61e6f-284">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-284">
        - TextBindings</span></span><br><span data-ttu-id="61e6f-285">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-285">
        - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="61e6f-286">Outlook</span><span class="sxs-lookup"><span data-stu-id="61e6f-286">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="61e6f-287">Plataforma</span><span class="sxs-lookup"><span data-stu-id="61e6f-287">Platform</span></span></th>
    <th><span data-ttu-id="61e6f-288">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="61e6f-288">Extension points</span></span></th>
    <th><span data-ttu-id="61e6f-289">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="61e6f-289">API requirement sets</span></span></th>
    <th><span data-ttu-id="61e6f-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="61e6f-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="61e6f-291">Office Online</span><span class="sxs-lookup"><span data-stu-id="61e6f-291">Office Online</span></span></td>
    <td> <span data-ttu-id="61e6f-292">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="61e6f-292">- Mail Read</span></span><br><span data-ttu-id="61e6f-293">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="61e6f-293">
      - Mail Compose</span></span><br><span data-ttu-id="61e6f-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="61e6f-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="61e6f-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="61e6f-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="61e6f-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="61e6f-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="61e6f-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="61e6f-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="61e6f-302">Não disponível</span><span class="sxs-lookup"><span data-stu-id="61e6f-302">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="61e6f-303">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="61e6f-303">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="61e6f-304">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="61e6f-304">- Mail Read</span></span><br><span data-ttu-id="61e6f-305">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="61e6f-305">
      - Mail Compose</span></span><br><span data-ttu-id="61e6f-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="61e6f-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="61e6f-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="61e6f-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="61e6f-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="61e6f-311">Não disponível</span><span class="sxs-lookup"><span data-stu-id="61e6f-311">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="61e6f-312">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="61e6f-312">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="61e6f-313">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="61e6f-313">- Mail Read</span></span><br><span data-ttu-id="61e6f-314">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="61e6f-314">
      - Mail Compose</span></span><br><span data-ttu-id="61e6f-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="61e6f-316">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="61e6f-316">
      - Modules</span></span></td>
    <td> <span data-ttu-id="61e6f-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="61e6f-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="61e6f-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="61e6f-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="61e6f-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="61e6f-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="61e6f-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="61e6f-324">Não disponível</span><span class="sxs-lookup"><span data-stu-id="61e6f-324">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="61e6f-325">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="61e6f-325">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="61e6f-326">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="61e6f-326">- Mail Read</span></span><br><span data-ttu-id="61e6f-327">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="61e6f-327">
      - Mail Compose</span></span><br><span data-ttu-id="61e6f-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="61e6f-329">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="61e6f-329">
      - Modules</span></span></td>
    <td> <span data-ttu-id="61e6f-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="61e6f-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="61e6f-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="61e6f-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="61e6f-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="61e6f-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="61e6f-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="61e6f-337">Não disponível</span><span class="sxs-lookup"><span data-stu-id="61e6f-337">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="61e6f-338">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="61e6f-338">Office for iOS</span></span></td>
    <td> <span data-ttu-id="61e6f-339">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="61e6f-339">- Mail Read</span></span><br><span data-ttu-id="61e6f-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="61e6f-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="61e6f-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="61e6f-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="61e6f-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="61e6f-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="61e6f-346">Não disponível</span><span class="sxs-lookup"><span data-stu-id="61e6f-346">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="61e6f-347">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="61e6f-347">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="61e6f-348">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="61e6f-348">- Mail Read</span></span><br><span data-ttu-id="61e6f-349">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="61e6f-349">
      - Mail Compose</span></span><br><span data-ttu-id="61e6f-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="61e6f-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="61e6f-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="61e6f-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="61e6f-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="61e6f-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="61e6f-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="61e6f-357">Não disponível</span><span class="sxs-lookup"><span data-stu-id="61e6f-357">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="61e6f-358">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="61e6f-358">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="61e6f-359">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="61e6f-359">- Mail Read</span></span><br><span data-ttu-id="61e6f-360">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="61e6f-360">
      - Mail Compose</span></span><br><span data-ttu-id="61e6f-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="61e6f-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="61e6f-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="61e6f-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="61e6f-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="61e6f-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="61e6f-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="61e6f-368">Não disponível</span><span class="sxs-lookup"><span data-stu-id="61e6f-368">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="61e6f-369">Office para Android</span><span class="sxs-lookup"><span data-stu-id="61e6f-369">Office for Android</span></span></td>
    <td> <span data-ttu-id="61e6f-370">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="61e6f-370">- Mail Read</span></span><br><span data-ttu-id="61e6f-371">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-371">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="61e6f-372">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-372">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="61e6f-373">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-373">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="61e6f-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="61e6f-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="61e6f-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="61e6f-377">Não disponível</span><span class="sxs-lookup"><span data-stu-id="61e6f-377">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="61e6f-378">Word</span><span class="sxs-lookup"><span data-stu-id="61e6f-378">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="61e6f-379">Plataforma</span><span class="sxs-lookup"><span data-stu-id="61e6f-379">Platform</span></span></th>
    <th><span data-ttu-id="61e6f-380">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="61e6f-380">Extension points</span></span></th>
    <th><span data-ttu-id="61e6f-381">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="61e6f-381">API requirement sets</span></span></th>
    <th><span data-ttu-id="61e6f-382"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="61e6f-382"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="61e6f-383">Office Online</span><span class="sxs-lookup"><span data-stu-id="61e6f-383">Office Online</span></span></td>
    <td> <span data-ttu-id="61e6f-384">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="61e6f-384">- TaskPane</span></span><br><span data-ttu-id="61e6f-385">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-385">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="61e6f-386">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-386">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="61e6f-387">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-387">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="61e6f-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="61e6f-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="61e6f-390">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="61e6f-390">- BindingEvents</span></span><br><span data-ttu-id="61e6f-391">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="61e6f-391">
         - CustomXmlParts</span></span><br><span data-ttu-id="61e6f-392">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="61e6f-392">
         - DocumentEvents</span></span><br><span data-ttu-id="61e6f-393">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="61e6f-393">
         - File</span></span><br><span data-ttu-id="61e6f-394">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-394">
         - HtmlCoercion</span></span><br><span data-ttu-id="61e6f-395">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-395">
         - ImageCoercion</span></span><br><span data-ttu-id="61e6f-396">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-396">
         - MatrixBindings</span></span><br><span data-ttu-id="61e6f-397">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-397">
         - MatrixCoercion</span></span><br><span data-ttu-id="61e6f-398">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-398">
         - OoxmlCoercion</span></span><br><span data-ttu-id="61e6f-399">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-399">
         - PdfFile</span></span><br><span data-ttu-id="61e6f-400">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="61e6f-400">
         - Selection</span></span><br><span data-ttu-id="61e6f-401">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="61e6f-401">
         - Settings</span></span><br><span data-ttu-id="61e6f-402">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-402">
         - TableBindings</span></span><br><span data-ttu-id="61e6f-403">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-403">
         - TableCoercion</span></span><br><span data-ttu-id="61e6f-404">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-404">
         - TextBindings</span></span><br><span data-ttu-id="61e6f-405">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-405">
         - TextCoercion</span></span><br><span data-ttu-id="61e6f-406">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-406">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="61e6f-407">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="61e6f-407">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="61e6f-408">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="61e6f-408">- TaskPane</span></span></td>
    <td> <span data-ttu-id="61e6f-409">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-409">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="61e6f-410">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="61e6f-410">- BindingEvents</span></span><br><span data-ttu-id="61e6f-411">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-411">
         - CompressedFile</span></span><br><span data-ttu-id="61e6f-412">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="61e6f-412">
         - CustomXmlParts</span></span><br><span data-ttu-id="61e6f-413">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="61e6f-413">
         - DocumentEvents</span></span><br><span data-ttu-id="61e6f-414">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="61e6f-414">
         - File</span></span><br><span data-ttu-id="61e6f-415">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-415">
         - HtmlCoercion</span></span><br><span data-ttu-id="61e6f-416">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-416">
         - ImageCoercion</span></span><br><span data-ttu-id="61e6f-417">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-417">
         - MatrixBindings</span></span><br><span data-ttu-id="61e6f-418">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-418">
         - MatrixCoercion</span></span><br><span data-ttu-id="61e6f-419">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-419">
         - OoxmlCoercion</span></span><br><span data-ttu-id="61e6f-420">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-420">
         - PdfFile</span></span><br><span data-ttu-id="61e6f-421">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="61e6f-421">
         - Selection</span></span><br><span data-ttu-id="61e6f-422">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="61e6f-422">
         - Settings</span></span><br><span data-ttu-id="61e6f-423">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-423">
         - TableBindings</span></span><br><span data-ttu-id="61e6f-424">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-424">
         - TableCoercion</span></span><br><span data-ttu-id="61e6f-425">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-425">
         - TextBindings</span></span><br><span data-ttu-id="61e6f-426">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-426">
         - TextCoercion</span></span><br><span data-ttu-id="61e6f-427">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-427">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="61e6f-428">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="61e6f-428">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="61e6f-429">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="61e6f-429">- TaskPane</span></span><br><span data-ttu-id="61e6f-430">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-430">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="61e6f-431">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-431">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="61e6f-432">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-432">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="61e6f-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="61e6f-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="61e6f-435">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="61e6f-435">- BindingEvents</span></span><br><span data-ttu-id="61e6f-436">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-436">
         - CompressedFile</span></span><br><span data-ttu-id="61e6f-437">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="61e6f-437">
         - CustomXmlParts</span></span><br><span data-ttu-id="61e6f-438">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="61e6f-438">
         - DocumentEvents</span></span><br><span data-ttu-id="61e6f-439">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="61e6f-439">
         - File</span></span><br><span data-ttu-id="61e6f-440">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-440">
         - HtmlCoercion</span></span><br><span data-ttu-id="61e6f-441">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-441">
         - ImageCoercion</span></span><br><span data-ttu-id="61e6f-442">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-442">
         - MatrixBindings</span></span><br><span data-ttu-id="61e6f-443">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-443">
         - MatrixCoercion</span></span><br><span data-ttu-id="61e6f-444">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-444">
         - OoxmlCoercion</span></span><br><span data-ttu-id="61e6f-445">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-445">
         - PdfFile</span></span><br><span data-ttu-id="61e6f-446">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="61e6f-446">
         - Selection</span></span><br><span data-ttu-id="61e6f-447">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="61e6f-447">
         - Settings</span></span><br><span data-ttu-id="61e6f-448">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-448">
         - TableBindings</span></span><br><span data-ttu-id="61e6f-449">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-449">
         - TableCoercion</span></span><br><span data-ttu-id="61e6f-450">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-450">
         - TextBindings</span></span><br><span data-ttu-id="61e6f-451">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-451">
         - TextCoercion</span></span><br><span data-ttu-id="61e6f-452">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-452">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="61e6f-453">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="61e6f-453">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="61e6f-454">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="61e6f-454">- TaskPane</span></span><br><span data-ttu-id="61e6f-455">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-455">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="61e6f-456">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-456">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="61e6f-457">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-457">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="61e6f-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="61e6f-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="61e6f-460">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="61e6f-460">- BindingEvents</span></span><br><span data-ttu-id="61e6f-461">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-461">
         - CompressedFile</span></span><br><span data-ttu-id="61e6f-462">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="61e6f-462">
         - CustomXmlParts</span></span><br><span data-ttu-id="61e6f-463">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="61e6f-463">
         - DocumentEvents</span></span><br><span data-ttu-id="61e6f-464">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="61e6f-464">
         - File</span></span><br><span data-ttu-id="61e6f-465">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-465">
         - HtmlCoercion</span></span><br><span data-ttu-id="61e6f-466">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-466">
         - ImageCoercion</span></span><br><span data-ttu-id="61e6f-467">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-467">
         - MatrixBindings</span></span><br><span data-ttu-id="61e6f-468">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-468">
         - MatrixCoercion</span></span><br><span data-ttu-id="61e6f-469">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-469">
         - OoxmlCoercion</span></span><br><span data-ttu-id="61e6f-470">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-470">
         - PdfFile</span></span><br><span data-ttu-id="61e6f-471">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="61e6f-471">
         - Selection</span></span><br><span data-ttu-id="61e6f-472">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="61e6f-472">
         - Settings</span></span><br><span data-ttu-id="61e6f-473">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-473">
         - TableBindings</span></span><br><span data-ttu-id="61e6f-474">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-474">
         - TableCoercion</span></span><br><span data-ttu-id="61e6f-475">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-475">
         - TextBindings</span></span><br><span data-ttu-id="61e6f-476">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-476">
         - TextCoercion</span></span><br><span data-ttu-id="61e6f-477">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-477">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="61e6f-478">Office para iPad</span><span class="sxs-lookup"><span data-stu-id="61e6f-478">Office for iPad</span></span></td>
    <td> <span data-ttu-id="61e6f-479">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="61e6f-479">- TaskPane</span></span></td>
    <td> <span data-ttu-id="61e6f-480">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-480">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="61e6f-481">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-481">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="61e6f-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="61e6f-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="61e6f-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="61e6f-484">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="61e6f-484">- BindingEvents</span></span><br><span data-ttu-id="61e6f-485">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-485">
         - CompressedFile</span></span><br><span data-ttu-id="61e6f-486">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="61e6f-486">
         - CustomXmlParts</span></span><br><span data-ttu-id="61e6f-487">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="61e6f-487">
         - DocumentEvents</span></span><br><span data-ttu-id="61e6f-488">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="61e6f-488">
         - File</span></span><br><span data-ttu-id="61e6f-489">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-489">
         - HtmlCoercion</span></span><br><span data-ttu-id="61e6f-490">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-490">
         - ImageCoercion</span></span><br><span data-ttu-id="61e6f-491">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-491">
         - MatrixBindings</span></span><br><span data-ttu-id="61e6f-492">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-492">
         - MatrixCoercion</span></span><br><span data-ttu-id="61e6f-493">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-493">
         - OoxmlCoercion</span></span><br><span data-ttu-id="61e6f-494">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-494">
         - PdfFile</span></span><br><span data-ttu-id="61e6f-495">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="61e6f-495">
         - Selection</span></span><br><span data-ttu-id="61e6f-496">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="61e6f-496">
         - Settings</span></span><br><span data-ttu-id="61e6f-497">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-497">
         - TableBindings</span></span><br><span data-ttu-id="61e6f-498">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-498">
         - TableCoercion</span></span><br><span data-ttu-id="61e6f-499">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-499">
         - TextBindings</span></span><br><span data-ttu-id="61e6f-500">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-500">
         - TextCoercion</span></span><br><span data-ttu-id="61e6f-501">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-501">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="61e6f-502">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="61e6f-502">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="61e6f-503">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="61e6f-503">- TaskPane</span></span><br><span data-ttu-id="61e6f-504">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-504">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="61e6f-505">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-505">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="61e6f-506">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-506">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="61e6f-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="61e6f-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="61e6f-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="61e6f-509">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="61e6f-509">- BindingEvents</span></span><br><span data-ttu-id="61e6f-510">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-510">
         - CompressedFile</span></span><br><span data-ttu-id="61e6f-511">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="61e6f-511">
         - CustomXmlParts</span></span><br><span data-ttu-id="61e6f-512">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="61e6f-512">
         - DocumentEvents</span></span><br><span data-ttu-id="61e6f-513">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="61e6f-513">
         - File</span></span><br><span data-ttu-id="61e6f-514">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-514">
         - HtmlCoercion</span></span><br><span data-ttu-id="61e6f-515">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-515">
         - ImageCoercion</span></span><br><span data-ttu-id="61e6f-516">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-516">
         - MatrixBindings</span></span><br><span data-ttu-id="61e6f-517">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-517">
         - MatrixCoercion</span></span><br><span data-ttu-id="61e6f-518">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-518">
         - OoxmlCoercion</span></span><br><span data-ttu-id="61e6f-519">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-519">
         - PdfFile</span></span><br><span data-ttu-id="61e6f-520">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="61e6f-520">
         - Selection</span></span><br><span data-ttu-id="61e6f-521">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="61e6f-521">
         - Settings</span></span><br><span data-ttu-id="61e6f-522">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-522">
         - TableBindings</span></span><br><span data-ttu-id="61e6f-523">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-523">
         - TableCoercion</span></span><br><span data-ttu-id="61e6f-524">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-524">
         - TextBindings</span></span><br><span data-ttu-id="61e6f-525">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-525">
         - TextCoercion</span></span><br><span data-ttu-id="61e6f-526">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-526">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="61e6f-527">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="61e6f-527">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="61e6f-528">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="61e6f-528">- TaskPane</span></span><br><span data-ttu-id="61e6f-529">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-529">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="61e6f-530">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-530">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="61e6f-531">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-531">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="61e6f-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="61e6f-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="61e6f-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="61e6f-534">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="61e6f-534">- BindingEvents</span></span><br><span data-ttu-id="61e6f-535">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-535">
         - CompressedFile</span></span><br><span data-ttu-id="61e6f-536">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="61e6f-536">
         - CustomXmlParts</span></span><br><span data-ttu-id="61e6f-537">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="61e6f-537">
         - DocumentEvents</span></span><br><span data-ttu-id="61e6f-538">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="61e6f-538">
         - File</span></span><br><span data-ttu-id="61e6f-539">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-539">
         - HtmlCoercion</span></span><br><span data-ttu-id="61e6f-540">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-540">
         - ImageCoercion</span></span><br><span data-ttu-id="61e6f-541">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-541">
         - MatrixBindings</span></span><br><span data-ttu-id="61e6f-542">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-542">
         - MatrixCoercion</span></span><br><span data-ttu-id="61e6f-543">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-543">
         - OoxmlCoercion</span></span><br><span data-ttu-id="61e6f-544">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-544">
         - PdfFile</span></span><br><span data-ttu-id="61e6f-545">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="61e6f-545">
         - Selection</span></span><br><span data-ttu-id="61e6f-546">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="61e6f-546">
         - Settings</span></span><br><span data-ttu-id="61e6f-547">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-547">
         - TableBindings</span></span><br><span data-ttu-id="61e6f-548">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-548">
         - TableCoercion</span></span><br><span data-ttu-id="61e6f-549">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="61e6f-549">
         - TextBindings</span></span><br><span data-ttu-id="61e6f-550">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-550">
         - TextCoercion</span></span><br><span data-ttu-id="61e6f-551">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-551">
         - TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="61e6f-552">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="61e6f-552">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="61e6f-553">Plataforma</span><span class="sxs-lookup"><span data-stu-id="61e6f-553">Platform</span></span></th>
    <th><span data-ttu-id="61e6f-554">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="61e6f-554">Extension points</span></span></th>
    <th><span data-ttu-id="61e6f-555">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="61e6f-555">API requirement sets</span></span></th>
    <th><span data-ttu-id="61e6f-556"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="61e6f-556"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="61e6f-557">Office Online</span><span class="sxs-lookup"><span data-stu-id="61e6f-557">Office Online</span></span></td>
    <td> <span data-ttu-id="61e6f-558">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="61e6f-558">- Content</span></span><br><span data-ttu-id="61e6f-559">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="61e6f-559">
         - TaskPane</span></span><br><span data-ttu-id="61e6f-560">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-560">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="61e6f-561">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-561">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="61e6f-562">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="61e6f-562">- ActiveView</span></span><br><span data-ttu-id="61e6f-563">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-563">
         - CompressedFile</span></span><br><span data-ttu-id="61e6f-564">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="61e6f-564">
         - DocumentEvents</span></span><br><span data-ttu-id="61e6f-565">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="61e6f-565">
         - File</span></span><br><span data-ttu-id="61e6f-566">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-566">
         - ImageCoercion</span></span><br><span data-ttu-id="61e6f-567">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-567">
         - PdfFile</span></span><br><span data-ttu-id="61e6f-568">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="61e6f-568">
         - Selection</span></span><br><span data-ttu-id="61e6f-569">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="61e6f-569">
         - Settings</span></span><br><span data-ttu-id="61e6f-570">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-570">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="61e6f-571">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="61e6f-571">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="61e6f-572">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="61e6f-572">- Content</span></span><br><span data-ttu-id="61e6f-573">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="61e6f-573">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="61e6f-574">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="61e6f-574">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="61e6f-575">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="61e6f-575">- ActiveView</span></span><br><span data-ttu-id="61e6f-576">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-576">
         - CompressedFile</span></span><br><span data-ttu-id="61e6f-577">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="61e6f-577">
         - DocumentEvents</span></span><br><span data-ttu-id="61e6f-578">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="61e6f-578">
         - File</span></span><br><span data-ttu-id="61e6f-579">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-579">
         - ImageCoercion</span></span><br><span data-ttu-id="61e6f-580">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-580">
         - PdfFile</span></span><br><span data-ttu-id="61e6f-581">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="61e6f-581">
         - Selection</span></span><br><span data-ttu-id="61e6f-582">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="61e6f-582">
         - Settings</span></span><br><span data-ttu-id="61e6f-583">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-583">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="61e6f-584">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="61e6f-584">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="61e6f-585">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="61e6f-585">- Content</span></span><br><span data-ttu-id="61e6f-586">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="61e6f-586">
         - TaskPane</span></span><br><span data-ttu-id="61e6f-587">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-587">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="61e6f-588">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-588">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="61e6f-589">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="61e6f-589">- ActiveView</span></span><br><span data-ttu-id="61e6f-590">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-590">
         - CompressedFile</span></span><br><span data-ttu-id="61e6f-591">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="61e6f-591">
         - DocumentEvents</span></span><br><span data-ttu-id="61e6f-592">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="61e6f-592">
         - File</span></span><br><span data-ttu-id="61e6f-593">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-593">
         - ImageCoercion</span></span><br><span data-ttu-id="61e6f-594">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-594">
         - PdfFile</span></span><br><span data-ttu-id="61e6f-595">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="61e6f-595">
         - Selection</span></span><br><span data-ttu-id="61e6f-596">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="61e6f-596">
         - Settings</span></span><br><span data-ttu-id="61e6f-597">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-597">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="61e6f-598">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="61e6f-598">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="61e6f-599">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="61e6f-599">- Content</span></span><br><span data-ttu-id="61e6f-600">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="61e6f-600">
         - TaskPane</span></span><br><span data-ttu-id="61e6f-601">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-601">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="61e6f-602">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-602">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="61e6f-603">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="61e6f-603">- ActiveView</span></span><br><span data-ttu-id="61e6f-604">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-604">
         - CompressedFile</span></span><br><span data-ttu-id="61e6f-605">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="61e6f-605">
         - DocumentEvents</span></span><br><span data-ttu-id="61e6f-606">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="61e6f-606">
         - File</span></span><br><span data-ttu-id="61e6f-607">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-607">
         - ImageCoercion</span></span><br><span data-ttu-id="61e6f-608">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-608">
         - PdfFile</span></span><br><span data-ttu-id="61e6f-609">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="61e6f-609">
         - Selection</span></span><br><span data-ttu-id="61e6f-610">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="61e6f-610">
         - Settings</span></span><br><span data-ttu-id="61e6f-611">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-611">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="61e6f-612">Office para iPad</span><span class="sxs-lookup"><span data-stu-id="61e6f-612">Office for iPad</span></span></td>
    <td> <span data-ttu-id="61e6f-613">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="61e6f-613">- Content</span></span><br><span data-ttu-id="61e6f-614">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="61e6f-614">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="61e6f-615">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-615">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="61e6f-616">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="61e6f-616">- ActiveView</span></span><br><span data-ttu-id="61e6f-617">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-617">
         - CompressedFile</span></span><br><span data-ttu-id="61e6f-618">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="61e6f-618">
         - DocumentEvents</span></span><br><span data-ttu-id="61e6f-619">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="61e6f-619">
         - File</span></span><br><span data-ttu-id="61e6f-620">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-620">
         - PdfFile</span></span><br><span data-ttu-id="61e6f-621">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="61e6f-621">
         - Selection</span></span><br><span data-ttu-id="61e6f-622">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="61e6f-622">
         - Settings</span></span><br><span data-ttu-id="61e6f-623">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-623">
         - TextCoercion</span></span><br><span data-ttu-id="61e6f-624">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-624">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="61e6f-625">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="61e6f-625">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="61e6f-626">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="61e6f-626">- Content</span></span><br><span data-ttu-id="61e6f-627">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="61e6f-627">
         - TaskPane</span></span><br><span data-ttu-id="61e6f-628">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-628">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="61e6f-629">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-629">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="61e6f-630">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="61e6f-630">- ActiveView</span></span><br><span data-ttu-id="61e6f-631">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-631">
         - CompressedFile</span></span><br><span data-ttu-id="61e6f-632">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="61e6f-632">
         - DocumentEvents</span></span><br><span data-ttu-id="61e6f-633">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="61e6f-633">
         - File</span></span><br><span data-ttu-id="61e6f-634">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-634">
         - ImageCoercion</span></span><br><span data-ttu-id="61e6f-635">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-635">
         - PdfFile</span></span><br><span data-ttu-id="61e6f-636">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="61e6f-636">
         - Selection</span></span><br><span data-ttu-id="61e6f-637">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="61e6f-637">
         - Settings</span></span><br><span data-ttu-id="61e6f-638">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-638">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="61e6f-639">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="61e6f-639">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="61e6f-640">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="61e6f-640">- Content</span></span><br><span data-ttu-id="61e6f-641">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="61e6f-641">
         - TaskPane</span></span><br><span data-ttu-id="61e6f-642">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-642">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="61e6f-643">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-643">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="61e6f-644">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="61e6f-644">- ActiveView</span></span><br><span data-ttu-id="61e6f-645">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-645">
         - CompressedFile</span></span><br><span data-ttu-id="61e6f-646">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="61e6f-646">
         - DocumentEvents</span></span><br><span data-ttu-id="61e6f-647">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="61e6f-647">
         - File</span></span><br><span data-ttu-id="61e6f-648">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-648">
         - ImageCoercion</span></span><br><span data-ttu-id="61e6f-649">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="61e6f-649">
         - PdfFile</span></span><br><span data-ttu-id="61e6f-650">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="61e6f-650">
         - Selection</span></span><br><span data-ttu-id="61e6f-651">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="61e6f-651">
         - Settings</span></span><br><span data-ttu-id="61e6f-652">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-652">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="61e6f-653">OneNote</span><span class="sxs-lookup"><span data-stu-id="61e6f-653">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="61e6f-654">Plataforma</span><span class="sxs-lookup"><span data-stu-id="61e6f-654">Platform</span></span></th>
    <th><span data-ttu-id="61e6f-655">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="61e6f-655">Extension points</span></span></th>
    <th><span data-ttu-id="61e6f-656">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="61e6f-656">API requirement sets</span></span></th>
    <th><span data-ttu-id="61e6f-657"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="61e6f-657"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="61e6f-658">Office Online</span><span class="sxs-lookup"><span data-stu-id="61e6f-658">Office Online</span></span></td>
    <td> <span data-ttu-id="61e6f-659">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="61e6f-659">- Content</span></span><br><span data-ttu-id="61e6f-660">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="61e6f-660">
         - TaskPane</span></span><br><span data-ttu-id="61e6f-661">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-661">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="61e6f-662">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-662">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="61e6f-663">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-663">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="61e6f-664">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="61e6f-664">- DocumentEvents</span></span><br><span data-ttu-id="61e6f-665">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-665">
         - HtmlCoercion</span></span><br><span data-ttu-id="61e6f-666">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-666">
         - ImageCoercion</span></span><br><span data-ttu-id="61e6f-667">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="61e6f-667">
         - Settings</span></span><br><span data-ttu-id="61e6f-668">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-668">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="61e6f-669">Project</span><span class="sxs-lookup"><span data-stu-id="61e6f-669">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="61e6f-670">Plataforma</span><span class="sxs-lookup"><span data-stu-id="61e6f-670">Platform</span></span></th>
    <th><span data-ttu-id="61e6f-671">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="61e6f-671">Extension points</span></span></th>
    <th><span data-ttu-id="61e6f-672">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="61e6f-672">API requirement sets</span></span></th>
    <th><span data-ttu-id="61e6f-673"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="61e6f-673"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="61e6f-674">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="61e6f-674">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="61e6f-675">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="61e6f-675">- TaskPane</span></span></td>
    <td> <span data-ttu-id="61e6f-676">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-676">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="61e6f-677">- Seleção</span><span class="sxs-lookup"><span data-stu-id="61e6f-677">- Selection</span></span><br><span data-ttu-id="61e6f-678">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-678">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="61e6f-679">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="61e6f-679">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="61e6f-680">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="61e6f-680">- TaskPane</span></span></td>
    <td> <span data-ttu-id="61e6f-681">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-681">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="61e6f-682">- Seleção</span><span class="sxs-lookup"><span data-stu-id="61e6f-682">- Selection</span></span><br><span data-ttu-id="61e6f-683">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-683">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="61e6f-684">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="61e6f-684">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="61e6f-685">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="61e6f-685">- TaskPane</span></span></td>
    <td> <span data-ttu-id="61e6f-686">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="61e6f-686">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="61e6f-687">- Seleção</span><span class="sxs-lookup"><span data-stu-id="61e6f-687">- Selection</span></span><br><span data-ttu-id="61e6f-688">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="61e6f-688">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="61e6f-689">Confira também</span><span class="sxs-lookup"><span data-stu-id="61e6f-689">See also</span></span>

- [<span data-ttu-id="61e6f-690">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="61e6f-690">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="61e6f-691">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="61e6f-691">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="61e6f-692">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="61e6f-692">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="61e6f-693">Referência da API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="61e6f-693">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
