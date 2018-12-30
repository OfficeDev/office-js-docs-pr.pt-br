---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos compatíveis com Excel, Word, Outlook, PowerPoint, OneNote e Project.
ms.date: 11/07/2018
ms.openlocfilehash: 9490fca9663737e2397de159169b545e3900289f
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/28/2018
ms.locfileid: "27458038"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="63fe3-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="63fe3-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="63fe3-104">Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="63fe3-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="63fe3-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que são compatíveis atualmente com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="63fe3-105">The following tables contain the available platforms, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="63fe3-p102">O número do build do Office 2016 instalado via MSI é 16.0.4266.1001. Esta versão só contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="63fe3-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="63fe3-108">Excel</span><span class="sxs-lookup"><span data-stu-id="63fe3-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="63fe3-109">Plataforma</span><span class="sxs-lookup"><span data-stu-id="63fe3-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="63fe3-110">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="63fe3-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="63fe3-111">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="63fe3-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="63fe3-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="63fe3-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="63fe3-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="63fe3-113">Office Online</span></span></td>
    <td> <span data-ttu-id="63fe3-114">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="63fe3-114">- TaskPane</span></span><br><span data-ttu-id="63fe3-115">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="63fe3-115">
        - Content</span></span><br><span data-ttu-id="63fe3-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="63fe3-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="63fe3-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="63fe3-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="63fe3-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="63fe3-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="63fe3-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="63fe3-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="63fe3-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="63fe3-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="63fe3-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="63fe3-126">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="63fe3-126">
        - BindingEvents</span></span><br><span data-ttu-id="63fe3-127">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-127">
        - CompressedFile</span></span><br><span data-ttu-id="63fe3-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="63fe3-128">
        - DocumentEvents</span></span><br><span data-ttu-id="63fe3-129">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="63fe3-129">
        - File</span></span><br><span data-ttu-id="63fe3-130">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-130">
        - MatrixBindings</span></span><br><span data-ttu-id="63fe3-131">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-131">
        - MatrixCoercion</span></span><br><span data-ttu-id="63fe3-132">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="63fe3-132">
        - Selection</span></span><br><span data-ttu-id="63fe3-133">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="63fe3-133">
        - Settings</span></span><br><span data-ttu-id="63fe3-134">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-134">
        - TableBindings</span></span><br><span data-ttu-id="63fe3-135">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-135">
        - TableCoercion</span></span><br><span data-ttu-id="63fe3-136">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-136">
        - TextBindings</span></span><br><span data-ttu-id="63fe3-137">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-137">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="63fe3-138">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="63fe3-138">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="63fe3-139">
        - TaskPane</span><span class="sxs-lookup"><span data-stu-id="63fe3-139">
        - TaskPane</span></span><br><span data-ttu-id="63fe3-140">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="63fe3-140">
        - Content</span></span></td>
    <td>  <span data-ttu-id="63fe3-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="63fe3-142">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="63fe3-142">
        - BindingEvents</span></span><br><span data-ttu-id="63fe3-143">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-143">
        - CompressedFile</span></span><br><span data-ttu-id="63fe3-144">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="63fe3-144">
        - DocumentEvents</span></span><br><span data-ttu-id="63fe3-145">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="63fe3-145">
        - File</span></span><br><span data-ttu-id="63fe3-146">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-146">
        - ImageCoercion</span></span><br><span data-ttu-id="63fe3-147">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-147">
        - MatrixBindings</span></span><br><span data-ttu-id="63fe3-148">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-148">
        - MatrixCoercion</span></span><br><span data-ttu-id="63fe3-149">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="63fe3-149">
        - Selection</span></span><br><span data-ttu-id="63fe3-150">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="63fe3-150">
        - Settings</span></span><br><span data-ttu-id="63fe3-151">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-151">
        - TableBindings</span></span><br><span data-ttu-id="63fe3-152">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-152">
        - TableCoercion</span></span><br><span data-ttu-id="63fe3-153">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-153">
        - TextBindings</span></span><br><span data-ttu-id="63fe3-154">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-154">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="63fe3-155">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="63fe3-155">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="63fe3-156">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="63fe3-156">- TaskPane</span></span><br><span data-ttu-id="63fe3-157">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="63fe3-157">
        - Content</span></span><br><span data-ttu-id="63fe3-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="63fe3-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="63fe3-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="63fe3-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="63fe3-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="63fe3-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="63fe3-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="63fe3-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="63fe3-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="63fe3-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="63fe3-168">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="63fe3-168">- BindingEvents</span></span><br><span data-ttu-id="63fe3-169">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-169">
        - CompressedFile</span></span><br><span data-ttu-id="63fe3-170">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="63fe3-170">
        - DocumentEvents</span></span><br><span data-ttu-id="63fe3-171">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="63fe3-171">
        - File</span></span><br><span data-ttu-id="63fe3-172">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-172">
        - ImageCoercion</span></span><br><span data-ttu-id="63fe3-173">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-173">
        - MatrixBindings</span></span><br><span data-ttu-id="63fe3-174">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-174">
        - MatrixCoercion</span></span><br><span data-ttu-id="63fe3-175">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="63fe3-175">
        - Selection</span></span><br><span data-ttu-id="63fe3-176">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="63fe3-176">
        - Settings</span></span><br><span data-ttu-id="63fe3-177">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-177">
        - TableBindings</span></span><br><span data-ttu-id="63fe3-178">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-178">
        - TableCoercion</span></span><br><span data-ttu-id="63fe3-179">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-179">
        - TextBindings</span></span><br><span data-ttu-id="63fe3-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-180">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="63fe3-181">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="63fe3-181">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="63fe3-182">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="63fe3-182">- TaskPane</span></span><br><span data-ttu-id="63fe3-183">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="63fe3-183">
        - Content</span></span><br><span data-ttu-id="63fe3-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="63fe3-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="63fe3-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="63fe3-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="63fe3-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="63fe3-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="63fe3-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="63fe3-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="63fe3-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="63fe3-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="63fe3-194">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="63fe3-194">- BindingEvents</span></span><br><span data-ttu-id="63fe3-195">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-195">
        - CompressedFile</span></span><br><span data-ttu-id="63fe3-196">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="63fe3-196">
        - DocumentEvents</span></span><br><span data-ttu-id="63fe3-197">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="63fe3-197">
        - File</span></span><br><span data-ttu-id="63fe3-198">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-198">
        - ImageCoercion</span></span><br><span data-ttu-id="63fe3-199">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-199">
        - MatrixBindings</span></span><br><span data-ttu-id="63fe3-200">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-200">
        - MatrixCoercion</span></span><br><span data-ttu-id="63fe3-201">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="63fe3-201">
        - Selection</span></span><br><span data-ttu-id="63fe3-202">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="63fe3-202">
        - Settings</span></span><br><span data-ttu-id="63fe3-203">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-203">
        - TableBindings</span></span><br><span data-ttu-id="63fe3-204">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-204">
        - TableCoercion</span></span><br><span data-ttu-id="63fe3-205">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-205">
        - TextBindings</span></span><br><span data-ttu-id="63fe3-206">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-206">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="63fe3-207">Office para iPad</span><span class="sxs-lookup"><span data-stu-id="63fe3-207">Office for iPad</span></span></td>
    <td><span data-ttu-id="63fe3-208">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="63fe3-208">- TaskPane</span></span><br><span data-ttu-id="63fe3-209">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="63fe3-209">
        - Content</span></span></td>
    <td><span data-ttu-id="63fe3-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="63fe3-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="63fe3-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="63fe3-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="63fe3-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="63fe3-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="63fe3-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="63fe3-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="63fe3-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="63fe3-219">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="63fe3-219">- BindingEvents</span></span><br><span data-ttu-id="63fe3-220">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-220">
        - CompressedFile</span></span><br><span data-ttu-id="63fe3-221">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="63fe3-221">
        - DocumentEvents</span></span><br><span data-ttu-id="63fe3-222">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="63fe3-222">
        - File</span></span><br><span data-ttu-id="63fe3-223">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-223">
        - ImageCoercion</span></span><br><span data-ttu-id="63fe3-224">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-224">
        - MatrixBindings</span></span><br><span data-ttu-id="63fe3-225">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-225">
        - MatrixCoercion</span></span><br><span data-ttu-id="63fe3-226">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="63fe3-226">
        - Selection</span></span><br><span data-ttu-id="63fe3-227">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="63fe3-227">
        - Settings</span></span><br><span data-ttu-id="63fe3-228">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-228">
        - TableBindings</span></span><br><span data-ttu-id="63fe3-229">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-229">
        - TableCoercion</span></span><br><span data-ttu-id="63fe3-230">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-230">
        - TextBindings</span></span><br><span data-ttu-id="63fe3-231">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-231">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="63fe3-232">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="63fe3-232">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="63fe3-233">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="63fe3-233">- TaskPane</span></span><br><span data-ttu-id="63fe3-234">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="63fe3-234">
        - Content</span></span><br><span data-ttu-id="63fe3-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="63fe3-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="63fe3-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="63fe3-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="63fe3-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="63fe3-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="63fe3-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="63fe3-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="63fe3-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="63fe3-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="63fe3-245">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="63fe3-245">- BindingEvents</span></span><br><span data-ttu-id="63fe3-246">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-246">
        - CompressedFile</span></span><br><span data-ttu-id="63fe3-247">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="63fe3-247">
        - DocumentEvents</span></span><br><span data-ttu-id="63fe3-248">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="63fe3-248">
        - File</span></span><br><span data-ttu-id="63fe3-249">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-249">
        - ImageCoercion</span></span><br><span data-ttu-id="63fe3-250">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-250">
        - MatrixBindings</span></span><br><span data-ttu-id="63fe3-251">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-251">
        - MatrixCoercion</span></span><br><span data-ttu-id="63fe3-252">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-252">
        - PdfFile</span></span><br><span data-ttu-id="63fe3-253">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="63fe3-253">
        - Selection</span></span><br><span data-ttu-id="63fe3-254">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="63fe3-254">
        - Settings</span></span><br><span data-ttu-id="63fe3-255">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-255">
        - TableBindings</span></span><br><span data-ttu-id="63fe3-256">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-256">
        - TableCoercion</span></span><br><span data-ttu-id="63fe3-257">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-257">
        - TextBindings</span></span><br><span data-ttu-id="63fe3-258">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-258">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="63fe3-259">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="63fe3-259">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="63fe3-260">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="63fe3-260">- TaskPane</span></span><br><span data-ttu-id="63fe3-261">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="63fe3-261">
        - Content</span></span><br><span data-ttu-id="63fe3-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="63fe3-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="63fe3-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="63fe3-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="63fe3-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="63fe3-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="63fe3-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="63fe3-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="63fe3-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="63fe3-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="63fe3-272">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="63fe3-272">- BindingEvents</span></span><br><span data-ttu-id="63fe3-273">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-273">
        - CompressedFile</span></span><br><span data-ttu-id="63fe3-274">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="63fe3-274">
        - DocumentEvents</span></span><br><span data-ttu-id="63fe3-275">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="63fe3-275">
        - File</span></span><br><span data-ttu-id="63fe3-276">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-276">
        - ImageCoercion</span></span><br><span data-ttu-id="63fe3-277">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-277">
        - MatrixBindings</span></span><br><span data-ttu-id="63fe3-278">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-278">
        - MatrixCoercion</span></span><br><span data-ttu-id="63fe3-279">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-279">
        - PdfFile</span></span><br><span data-ttu-id="63fe3-280">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="63fe3-280">
        - Selection</span></span><br><span data-ttu-id="63fe3-281">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="63fe3-281">
        - Settings</span></span><br><span data-ttu-id="63fe3-282">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-282">
        - TableBindings</span></span><br><span data-ttu-id="63fe3-283">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-283">
        - TableCoercion</span></span><br><span data-ttu-id="63fe3-284">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-284">
        - TextBindings</span></span><br><span data-ttu-id="63fe3-285">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-285">
        - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="63fe3-286">Outlook</span><span class="sxs-lookup"><span data-stu-id="63fe3-286">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="63fe3-287">Plataforma</span><span class="sxs-lookup"><span data-stu-id="63fe3-287">Platform</span></span></th>
    <th><span data-ttu-id="63fe3-288">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="63fe3-288">Extension points</span></span></th>
    <th><span data-ttu-id="63fe3-289">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="63fe3-289">API requirement sets</span></span></th>
    <th><span data-ttu-id="63fe3-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="63fe3-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="63fe3-291">Office Online</span><span class="sxs-lookup"><span data-stu-id="63fe3-291">Office Online</span></span></td>
    <td> <span data-ttu-id="63fe3-292">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="63fe3-292">- Mail Read</span></span><br><span data-ttu-id="63fe3-293">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="63fe3-293">
      - Mail Compose</span></span><br><span data-ttu-id="63fe3-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="63fe3-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="63fe3-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="63fe3-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="63fe3-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="63fe3-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="63fe3-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="63fe3-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="63fe3-302">Não disponível</span><span class="sxs-lookup"><span data-stu-id="63fe3-302">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="63fe3-303">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="63fe3-303">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="63fe3-304">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="63fe3-304">- Mail Read</span></span><br><span data-ttu-id="63fe3-305">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="63fe3-305">
      - Mail Compose</span></span><br><span data-ttu-id="63fe3-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="63fe3-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="63fe3-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="63fe3-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="63fe3-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="63fe3-311">Não disponível</span><span class="sxs-lookup"><span data-stu-id="63fe3-311">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="63fe3-312">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="63fe3-312">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="63fe3-313">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="63fe3-313">- Mail Read</span></span><br><span data-ttu-id="63fe3-314">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="63fe3-314">
      - Mail Compose</span></span><br><span data-ttu-id="63fe3-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="63fe3-316">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="63fe3-316">
      - Modules</span></span></td>
    <td> <span data-ttu-id="63fe3-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="63fe3-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="63fe3-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="63fe3-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="63fe3-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="63fe3-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="63fe3-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="63fe3-324">Não disponível</span><span class="sxs-lookup"><span data-stu-id="63fe3-324">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="63fe3-325">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="63fe3-325">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="63fe3-326">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="63fe3-326">- Mail Read</span></span><br><span data-ttu-id="63fe3-327">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="63fe3-327">
      - Mail Compose</span></span><br><span data-ttu-id="63fe3-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="63fe3-329">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="63fe3-329">
      - Modules</span></span></td>
    <td> <span data-ttu-id="63fe3-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="63fe3-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="63fe3-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="63fe3-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="63fe3-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="63fe3-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="63fe3-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="63fe3-337">Não disponível</span><span class="sxs-lookup"><span data-stu-id="63fe3-337">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="63fe3-338">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="63fe3-338">Office for iOS</span></span></td>
    <td> <span data-ttu-id="63fe3-339">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="63fe3-339">- Mail Read</span></span><br><span data-ttu-id="63fe3-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="63fe3-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="63fe3-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="63fe3-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="63fe3-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="63fe3-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="63fe3-346">Não disponível</span><span class="sxs-lookup"><span data-stu-id="63fe3-346">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="63fe3-347">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="63fe3-347">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="63fe3-348">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="63fe3-348">- Mail Read</span></span><br><span data-ttu-id="63fe3-349">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="63fe3-349">
      - Mail Compose</span></span><br><span data-ttu-id="63fe3-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="63fe3-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="63fe3-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="63fe3-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="63fe3-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="63fe3-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="63fe3-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="63fe3-357">Não disponível</span><span class="sxs-lookup"><span data-stu-id="63fe3-357">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="63fe3-358">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="63fe3-358">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="63fe3-359">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="63fe3-359">- Mail Read</span></span><br><span data-ttu-id="63fe3-360">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="63fe3-360">
      - Mail Compose</span></span><br><span data-ttu-id="63fe3-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="63fe3-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="63fe3-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="63fe3-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="63fe3-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="63fe3-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="63fe3-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="63fe3-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="63fe3-369">Não disponível</span><span class="sxs-lookup"><span data-stu-id="63fe3-369">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="63fe3-370">Office para Android</span><span class="sxs-lookup"><span data-stu-id="63fe3-370">Office for Android</span></span></td>
    <td> <span data-ttu-id="63fe3-371">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="63fe3-371">- Mail Read</span></span><br><span data-ttu-id="63fe3-372">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-372">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="63fe3-373">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-373">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="63fe3-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="63fe3-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="63fe3-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="63fe3-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="63fe3-378">Não disponível</span><span class="sxs-lookup"><span data-stu-id="63fe3-378">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="63fe3-379">Word</span><span class="sxs-lookup"><span data-stu-id="63fe3-379">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="63fe3-380">Plataforma</span><span class="sxs-lookup"><span data-stu-id="63fe3-380">Platform</span></span></th>
    <th><span data-ttu-id="63fe3-381">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="63fe3-381">Extension points</span></span></th>
    <th><span data-ttu-id="63fe3-382">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="63fe3-382">API requirement sets</span></span></th>
    <th><span data-ttu-id="63fe3-383"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="63fe3-383"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="63fe3-384">Office Online</span><span class="sxs-lookup"><span data-stu-id="63fe3-384">Office Online</span></span></td>
    <td> <span data-ttu-id="63fe3-385">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="63fe3-385">- TaskPane</span></span><br><span data-ttu-id="63fe3-386">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-386">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="63fe3-387">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-387">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="63fe3-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="63fe3-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="63fe3-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="63fe3-391">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="63fe3-391">- BindingEvents</span></span><br><span data-ttu-id="63fe3-392">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="63fe3-392">
         - CustomXmlParts</span></span><br><span data-ttu-id="63fe3-393">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="63fe3-393">
         - DocumentEvents</span></span><br><span data-ttu-id="63fe3-394">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="63fe3-394">
         - File</span></span><br><span data-ttu-id="63fe3-395">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-395">
         - HtmlCoercion</span></span><br><span data-ttu-id="63fe3-396">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-396">
         - ImageCoercion</span></span><br><span data-ttu-id="63fe3-397">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-397">
         - MatrixBindings</span></span><br><span data-ttu-id="63fe3-398">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-398">
         - MatrixCoercion</span></span><br><span data-ttu-id="63fe3-399">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-399">
         - OoxmlCoercion</span></span><br><span data-ttu-id="63fe3-400">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-400">
         - PdfFile</span></span><br><span data-ttu-id="63fe3-401">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="63fe3-401">
         - Selection</span></span><br><span data-ttu-id="63fe3-402">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="63fe3-402">
         - Settings</span></span><br><span data-ttu-id="63fe3-403">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-403">
         - TableBindings</span></span><br><span data-ttu-id="63fe3-404">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-404">
         - TableCoercion</span></span><br><span data-ttu-id="63fe3-405">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-405">
         - TextBindings</span></span><br><span data-ttu-id="63fe3-406">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-406">
         - TextCoercion</span></span><br><span data-ttu-id="63fe3-407">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-407">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="63fe3-408">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="63fe3-408">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="63fe3-409">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="63fe3-409">- TaskPane</span></span></td>
    <td> <span data-ttu-id="63fe3-410">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-410">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="63fe3-411">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="63fe3-411">- BindingEvents</span></span><br><span data-ttu-id="63fe3-412">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-412">
         - CompressedFile</span></span><br><span data-ttu-id="63fe3-413">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="63fe3-413">
         - CustomXmlParts</span></span><br><span data-ttu-id="63fe3-414">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="63fe3-414">
         - DocumentEvents</span></span><br><span data-ttu-id="63fe3-415">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="63fe3-415">
         - File</span></span><br><span data-ttu-id="63fe3-416">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-416">
         - HtmlCoercion</span></span><br><span data-ttu-id="63fe3-417">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-417">
         - ImageCoercion</span></span><br><span data-ttu-id="63fe3-418">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-418">
         - MatrixBindings</span></span><br><span data-ttu-id="63fe3-419">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-419">
         - MatrixCoercion</span></span><br><span data-ttu-id="63fe3-420">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-420">
         - OoxmlCoercion</span></span><br><span data-ttu-id="63fe3-421">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-421">
         - PdfFile</span></span><br><span data-ttu-id="63fe3-422">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="63fe3-422">
         - Selection</span></span><br><span data-ttu-id="63fe3-423">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="63fe3-423">
         - Settings</span></span><br><span data-ttu-id="63fe3-424">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-424">
         - TableBindings</span></span><br><span data-ttu-id="63fe3-425">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-425">
         - TableCoercion</span></span><br><span data-ttu-id="63fe3-426">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-426">
         - TextBindings</span></span><br><span data-ttu-id="63fe3-427">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-427">
         - TextCoercion</span></span><br><span data-ttu-id="63fe3-428">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-428">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="63fe3-429">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="63fe3-429">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="63fe3-430">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="63fe3-430">- TaskPane</span></span><br><span data-ttu-id="63fe3-431">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-431">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="63fe3-432">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-432">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="63fe3-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="63fe3-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="63fe3-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="63fe3-436">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="63fe3-436">- BindingEvents</span></span><br><span data-ttu-id="63fe3-437">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-437">
         - CompressedFile</span></span><br><span data-ttu-id="63fe3-438">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="63fe3-438">
         - CustomXmlParts</span></span><br><span data-ttu-id="63fe3-439">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="63fe3-439">
         - DocumentEvents</span></span><br><span data-ttu-id="63fe3-440">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="63fe3-440">
         - File</span></span><br><span data-ttu-id="63fe3-441">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-441">
         - HtmlCoercion</span></span><br><span data-ttu-id="63fe3-442">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-442">
         - ImageCoercion</span></span><br><span data-ttu-id="63fe3-443">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-443">
         - MatrixBindings</span></span><br><span data-ttu-id="63fe3-444">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-444">
         - MatrixCoercion</span></span><br><span data-ttu-id="63fe3-445">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-445">
         - OoxmlCoercion</span></span><br><span data-ttu-id="63fe3-446">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-446">
         - PdfFile</span></span><br><span data-ttu-id="63fe3-447">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="63fe3-447">
         - Selection</span></span><br><span data-ttu-id="63fe3-448">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="63fe3-448">
         - Settings</span></span><br><span data-ttu-id="63fe3-449">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-449">
         - TableBindings</span></span><br><span data-ttu-id="63fe3-450">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-450">
         - TableCoercion</span></span><br><span data-ttu-id="63fe3-451">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-451">
         - TextBindings</span></span><br><span data-ttu-id="63fe3-452">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-452">
         - TextCoercion</span></span><br><span data-ttu-id="63fe3-453">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-453">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="63fe3-454">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="63fe3-454">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="63fe3-455">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="63fe3-455">- TaskPane</span></span><br><span data-ttu-id="63fe3-456">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-456">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="63fe3-457">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-457">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="63fe3-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="63fe3-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="63fe3-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="63fe3-461">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="63fe3-461">- BindingEvents</span></span><br><span data-ttu-id="63fe3-462">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-462">
         - CompressedFile</span></span><br><span data-ttu-id="63fe3-463">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="63fe3-463">
         - CustomXmlParts</span></span><br><span data-ttu-id="63fe3-464">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="63fe3-464">
         - DocumentEvents</span></span><br><span data-ttu-id="63fe3-465">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="63fe3-465">
         - File</span></span><br><span data-ttu-id="63fe3-466">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-466">
         - HtmlCoercion</span></span><br><span data-ttu-id="63fe3-467">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-467">
         - ImageCoercion</span></span><br><span data-ttu-id="63fe3-468">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-468">
         - MatrixBindings</span></span><br><span data-ttu-id="63fe3-469">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-469">
         - MatrixCoercion</span></span><br><span data-ttu-id="63fe3-470">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-470">
         - OoxmlCoercion</span></span><br><span data-ttu-id="63fe3-471">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-471">
         - PdfFile</span></span><br><span data-ttu-id="63fe3-472">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="63fe3-472">
         - Selection</span></span><br><span data-ttu-id="63fe3-473">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="63fe3-473">
         - Settings</span></span><br><span data-ttu-id="63fe3-474">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-474">
         - TableBindings</span></span><br><span data-ttu-id="63fe3-475">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-475">
         - TableCoercion</span></span><br><span data-ttu-id="63fe3-476">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-476">
         - TextBindings</span></span><br><span data-ttu-id="63fe3-477">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-477">
         - TextCoercion</span></span><br><span data-ttu-id="63fe3-478">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-478">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="63fe3-479">Office para iPad</span><span class="sxs-lookup"><span data-stu-id="63fe3-479">Office for iPad</span></span></td>
    <td> <span data-ttu-id="63fe3-480">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="63fe3-480">- TaskPane</span></span></td>
    <td> <span data-ttu-id="63fe3-481">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-481">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="63fe3-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="63fe3-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="63fe3-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="63fe3-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="63fe3-485">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="63fe3-485">- BindingEvents</span></span><br><span data-ttu-id="63fe3-486">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-486">
         - CompressedFile</span></span><br><span data-ttu-id="63fe3-487">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="63fe3-487">
         - CustomXmlParts</span></span><br><span data-ttu-id="63fe3-488">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="63fe3-488">
         - DocumentEvents</span></span><br><span data-ttu-id="63fe3-489">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="63fe3-489">
         - File</span></span><br><span data-ttu-id="63fe3-490">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-490">
         - HtmlCoercion</span></span><br><span data-ttu-id="63fe3-491">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-491">
         - ImageCoercion</span></span><br><span data-ttu-id="63fe3-492">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-492">
         - MatrixBindings</span></span><br><span data-ttu-id="63fe3-493">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-493">
         - MatrixCoercion</span></span><br><span data-ttu-id="63fe3-494">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-494">
         - OoxmlCoercion</span></span><br><span data-ttu-id="63fe3-495">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-495">
         - PdfFile</span></span><br><span data-ttu-id="63fe3-496">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="63fe3-496">
         - Selection</span></span><br><span data-ttu-id="63fe3-497">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="63fe3-497">
         - Settings</span></span><br><span data-ttu-id="63fe3-498">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-498">
         - TableBindings</span></span><br><span data-ttu-id="63fe3-499">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-499">
         - TableCoercion</span></span><br><span data-ttu-id="63fe3-500">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-500">
         - TextBindings</span></span><br><span data-ttu-id="63fe3-501">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-501">
         - TextCoercion</span></span><br><span data-ttu-id="63fe3-502">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-502">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="63fe3-503">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="63fe3-503">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="63fe3-504">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="63fe3-504">- TaskPane</span></span><br><span data-ttu-id="63fe3-505">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-505">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="63fe3-506">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-506">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="63fe3-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="63fe3-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="63fe3-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="63fe3-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="63fe3-510">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="63fe3-510">- BindingEvents</span></span><br><span data-ttu-id="63fe3-511">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-511">
         - CompressedFile</span></span><br><span data-ttu-id="63fe3-512">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="63fe3-512">
         - CustomXmlParts</span></span><br><span data-ttu-id="63fe3-513">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="63fe3-513">
         - DocumentEvents</span></span><br><span data-ttu-id="63fe3-514">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="63fe3-514">
         - File</span></span><br><span data-ttu-id="63fe3-515">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-515">
         - HtmlCoercion</span></span><br><span data-ttu-id="63fe3-516">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-516">
         - ImageCoercion</span></span><br><span data-ttu-id="63fe3-517">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-517">
         - MatrixBindings</span></span><br><span data-ttu-id="63fe3-518">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-518">
         - MatrixCoercion</span></span><br><span data-ttu-id="63fe3-519">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-519">
         - OoxmlCoercion</span></span><br><span data-ttu-id="63fe3-520">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-520">
         - PdfFile</span></span><br><span data-ttu-id="63fe3-521">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="63fe3-521">
         - Selection</span></span><br><span data-ttu-id="63fe3-522">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="63fe3-522">
         - Settings</span></span><br><span data-ttu-id="63fe3-523">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-523">
         - TableBindings</span></span><br><span data-ttu-id="63fe3-524">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-524">
         - TableCoercion</span></span><br><span data-ttu-id="63fe3-525">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-525">
         - TextBindings</span></span><br><span data-ttu-id="63fe3-526">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-526">
         - TextCoercion</span></span><br><span data-ttu-id="63fe3-527">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-527">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="63fe3-528">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="63fe3-528">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="63fe3-529">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="63fe3-529">- TaskPane</span></span><br><span data-ttu-id="63fe3-530">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-530">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="63fe3-531">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-531">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="63fe3-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="63fe3-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="63fe3-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="63fe3-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="63fe3-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="63fe3-535">- BindingEvents</span></span><br><span data-ttu-id="63fe3-536">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-536">
         - CompressedFile</span></span><br><span data-ttu-id="63fe3-537">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="63fe3-537">
         - CustomXmlParts</span></span><br><span data-ttu-id="63fe3-538">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="63fe3-538">
         - DocumentEvents</span></span><br><span data-ttu-id="63fe3-539">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="63fe3-539">
         - File</span></span><br><span data-ttu-id="63fe3-540">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-540">
         - HtmlCoercion</span></span><br><span data-ttu-id="63fe3-541">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-541">
         - ImageCoercion</span></span><br><span data-ttu-id="63fe3-542">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-542">
         - MatrixBindings</span></span><br><span data-ttu-id="63fe3-543">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-543">
         - MatrixCoercion</span></span><br><span data-ttu-id="63fe3-544">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-544">
         - OoxmlCoercion</span></span><br><span data-ttu-id="63fe3-545">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-545">
         - PdfFile</span></span><br><span data-ttu-id="63fe3-546">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="63fe3-546">
         - Selection</span></span><br><span data-ttu-id="63fe3-547">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="63fe3-547">
         - Settings</span></span><br><span data-ttu-id="63fe3-548">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-548">
         - TableBindings</span></span><br><span data-ttu-id="63fe3-549">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-549">
         - TableCoercion</span></span><br><span data-ttu-id="63fe3-550">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="63fe3-550">
         - TextBindings</span></span><br><span data-ttu-id="63fe3-551">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-551">
         - TextCoercion</span></span><br><span data-ttu-id="63fe3-552">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-552">
         - TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="63fe3-553">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="63fe3-553">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="63fe3-554">Plataforma</span><span class="sxs-lookup"><span data-stu-id="63fe3-554">Platform</span></span></th>
    <th><span data-ttu-id="63fe3-555">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="63fe3-555">Extension points</span></span></th>
    <th><span data-ttu-id="63fe3-556">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="63fe3-556">API requirement sets</span></span></th>
    <th><span data-ttu-id="63fe3-557"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="63fe3-557"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="63fe3-558">Office Online</span><span class="sxs-lookup"><span data-stu-id="63fe3-558">Office Online</span></span></td>
    <td> <span data-ttu-id="63fe3-559">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="63fe3-559">- Content</span></span><br><span data-ttu-id="63fe3-560">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="63fe3-560">
         - TaskPane</span></span><br><span data-ttu-id="63fe3-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="63fe3-562">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-562">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="63fe3-563">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="63fe3-563">- ActiveView</span></span><br><span data-ttu-id="63fe3-564">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-564">
         - CompressedFile</span></span><br><span data-ttu-id="63fe3-565">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="63fe3-565">
         - DocumentEvents</span></span><br><span data-ttu-id="63fe3-566">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="63fe3-566">
         - File</span></span><br><span data-ttu-id="63fe3-567">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-567">
         - ImageCoercion</span></span><br><span data-ttu-id="63fe3-568">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-568">
         - PdfFile</span></span><br><span data-ttu-id="63fe3-569">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="63fe3-569">
         - Selection</span></span><br><span data-ttu-id="63fe3-570">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="63fe3-570">
         - Settings</span></span><br><span data-ttu-id="63fe3-571">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-571">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="63fe3-572">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="63fe3-572">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="63fe3-573">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="63fe3-573">- Content</span></span><br><span data-ttu-id="63fe3-574">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="63fe3-574">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="63fe3-575">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="63fe3-575">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="63fe3-576">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="63fe3-576">- ActiveView</span></span><br><span data-ttu-id="63fe3-577">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-577">
         - CompressedFile</span></span><br><span data-ttu-id="63fe3-578">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="63fe3-578">
         - DocumentEvents</span></span><br><span data-ttu-id="63fe3-579">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="63fe3-579">
         - File</span></span><br><span data-ttu-id="63fe3-580">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-580">
         - ImageCoercion</span></span><br><span data-ttu-id="63fe3-581">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-581">
         - PdfFile</span></span><br><span data-ttu-id="63fe3-582">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="63fe3-582">
         - Selection</span></span><br><span data-ttu-id="63fe3-583">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="63fe3-583">
         - Settings</span></span><br><span data-ttu-id="63fe3-584">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-584">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="63fe3-585">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="63fe3-585">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="63fe3-586">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="63fe3-586">- Content</span></span><br><span data-ttu-id="63fe3-587">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="63fe3-587">
         - TaskPane</span></span><br><span data-ttu-id="63fe3-588">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-588">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="63fe3-589">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-589">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="63fe3-590">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="63fe3-590">- ActiveView</span></span><br><span data-ttu-id="63fe3-591">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-591">
         - CompressedFile</span></span><br><span data-ttu-id="63fe3-592">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="63fe3-592">
         - DocumentEvents</span></span><br><span data-ttu-id="63fe3-593">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="63fe3-593">
         - File</span></span><br><span data-ttu-id="63fe3-594">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-594">
         - ImageCoercion</span></span><br><span data-ttu-id="63fe3-595">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-595">
         - PdfFile</span></span><br><span data-ttu-id="63fe3-596">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="63fe3-596">
         - Selection</span></span><br><span data-ttu-id="63fe3-597">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="63fe3-597">
         - Settings</span></span><br><span data-ttu-id="63fe3-598">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-598">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="63fe3-599">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="63fe3-599">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="63fe3-600">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="63fe3-600">- Content</span></span><br><span data-ttu-id="63fe3-601">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="63fe3-601">
         - TaskPane</span></span><br><span data-ttu-id="63fe3-602">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-602">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="63fe3-603">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-603">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="63fe3-604">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="63fe3-604">- ActiveView</span></span><br><span data-ttu-id="63fe3-605">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-605">
         - CompressedFile</span></span><br><span data-ttu-id="63fe3-606">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="63fe3-606">
         - DocumentEvents</span></span><br><span data-ttu-id="63fe3-607">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="63fe3-607">
         - File</span></span><br><span data-ttu-id="63fe3-608">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-608">
         - ImageCoercion</span></span><br><span data-ttu-id="63fe3-609">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-609">
         - PdfFile</span></span><br><span data-ttu-id="63fe3-610">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="63fe3-610">
         - Selection</span></span><br><span data-ttu-id="63fe3-611">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="63fe3-611">
         - Settings</span></span><br><span data-ttu-id="63fe3-612">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-612">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="63fe3-613">Office para iPad</span><span class="sxs-lookup"><span data-stu-id="63fe3-613">Office for iPad</span></span></td>
    <td> <span data-ttu-id="63fe3-614">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="63fe3-614">- Content</span></span><br><span data-ttu-id="63fe3-615">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="63fe3-615">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="63fe3-616">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-616">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="63fe3-617">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="63fe3-617">- ActiveView</span></span><br><span data-ttu-id="63fe3-618">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-618">
         - CompressedFile</span></span><br><span data-ttu-id="63fe3-619">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="63fe3-619">
         - DocumentEvents</span></span><br><span data-ttu-id="63fe3-620">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="63fe3-620">
         - File</span></span><br><span data-ttu-id="63fe3-621">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-621">
         - PdfFile</span></span><br><span data-ttu-id="63fe3-622">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="63fe3-622">
         - Selection</span></span><br><span data-ttu-id="63fe3-623">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="63fe3-623">
         - Settings</span></span><br><span data-ttu-id="63fe3-624">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-624">
         - TextCoercion</span></span><br><span data-ttu-id="63fe3-625">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-625">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="63fe3-626">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="63fe3-626">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="63fe3-627">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="63fe3-627">- Content</span></span><br><span data-ttu-id="63fe3-628">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="63fe3-628">
         - TaskPane</span></span><br><span data-ttu-id="63fe3-629">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-629">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="63fe3-630">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-630">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="63fe3-631">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="63fe3-631">- ActiveView</span></span><br><span data-ttu-id="63fe3-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-632">
         - CompressedFile</span></span><br><span data-ttu-id="63fe3-633">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="63fe3-633">
         - DocumentEvents</span></span><br><span data-ttu-id="63fe3-634">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="63fe3-634">
         - File</span></span><br><span data-ttu-id="63fe3-635">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-635">
         - ImageCoercion</span></span><br><span data-ttu-id="63fe3-636">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-636">
         - PdfFile</span></span><br><span data-ttu-id="63fe3-637">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="63fe3-637">
         - Selection</span></span><br><span data-ttu-id="63fe3-638">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="63fe3-638">
         - Settings</span></span><br><span data-ttu-id="63fe3-639">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-639">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="63fe3-640">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="63fe3-640">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="63fe3-641">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="63fe3-641">- Content</span></span><br><span data-ttu-id="63fe3-642">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="63fe3-642">
         - TaskPane</span></span><br><span data-ttu-id="63fe3-643">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-643">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="63fe3-644">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-644">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="63fe3-645">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="63fe3-645">- ActiveView</span></span><br><span data-ttu-id="63fe3-646">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-646">
         - CompressedFile</span></span><br><span data-ttu-id="63fe3-647">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="63fe3-647">
         - DocumentEvents</span></span><br><span data-ttu-id="63fe3-648">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="63fe3-648">
         - File</span></span><br><span data-ttu-id="63fe3-649">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-649">
         - ImageCoercion</span></span><br><span data-ttu-id="63fe3-650">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="63fe3-650">
         - PdfFile</span></span><br><span data-ttu-id="63fe3-651">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="63fe3-651">
         - Selection</span></span><br><span data-ttu-id="63fe3-652">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="63fe3-652">
         - Settings</span></span><br><span data-ttu-id="63fe3-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-653">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="63fe3-654">OneNote</span><span class="sxs-lookup"><span data-stu-id="63fe3-654">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="63fe3-655">Plataforma</span><span class="sxs-lookup"><span data-stu-id="63fe3-655">Platform</span></span></th>
    <th><span data-ttu-id="63fe3-656">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="63fe3-656">Extension points</span></span></th>
    <th><span data-ttu-id="63fe3-657">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="63fe3-657">API requirement sets</span></span></th>
    <th><span data-ttu-id="63fe3-658"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="63fe3-658"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="63fe3-659">Office Online</span><span class="sxs-lookup"><span data-stu-id="63fe3-659">Office Online</span></span></td>
    <td> <span data-ttu-id="63fe3-660">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="63fe3-660">- Content</span></span><br><span data-ttu-id="63fe3-661">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="63fe3-661">
         - TaskPane</span></span><br><span data-ttu-id="63fe3-662">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-662">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="63fe3-663">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-663">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="63fe3-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="63fe3-665">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="63fe3-665">- DocumentEvents</span></span><br><span data-ttu-id="63fe3-666">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-666">
         - HtmlCoercion</span></span><br><span data-ttu-id="63fe3-667">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-667">
         - ImageCoercion</span></span><br><span data-ttu-id="63fe3-668">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="63fe3-668">
         - Settings</span></span><br><span data-ttu-id="63fe3-669">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-669">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="63fe3-670">Project</span><span class="sxs-lookup"><span data-stu-id="63fe3-670">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="63fe3-671">Plataforma</span><span class="sxs-lookup"><span data-stu-id="63fe3-671">Platform</span></span></th>
    <th><span data-ttu-id="63fe3-672">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="63fe3-672">Extension points</span></span></th>
    <th><span data-ttu-id="63fe3-673">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="63fe3-673">API requirement sets</span></span></th>
    <th><span data-ttu-id="63fe3-674"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="63fe3-674"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="63fe3-675">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="63fe3-675">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="63fe3-676">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="63fe3-676">- TaskPane</span></span></td>
    <td> <span data-ttu-id="63fe3-677">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-677">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="63fe3-678">- Seleção</span><span class="sxs-lookup"><span data-stu-id="63fe3-678">- Selection</span></span><br><span data-ttu-id="63fe3-679">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-679">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="63fe3-680">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="63fe3-680">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="63fe3-681">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="63fe3-681">- TaskPane</span></span></td>
    <td> <span data-ttu-id="63fe3-682">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-682">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="63fe3-683">- Seleção</span><span class="sxs-lookup"><span data-stu-id="63fe3-683">- Selection</span></span><br><span data-ttu-id="63fe3-684">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-684">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="63fe3-685">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="63fe3-685">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="63fe3-686">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="63fe3-686">- TaskPane</span></span></td>
    <td> <span data-ttu-id="63fe3-687">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="63fe3-687">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="63fe3-688">- Seleção</span><span class="sxs-lookup"><span data-stu-id="63fe3-688">- Selection</span></span><br><span data-ttu-id="63fe3-689">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="63fe3-689">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="63fe3-690">Confira também</span><span class="sxs-lookup"><span data-stu-id="63fe3-690">See also</span></span>

- [<span data-ttu-id="63fe3-691">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="63fe3-691">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="63fe3-692">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="63fe3-692">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="63fe3-693">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="63fe3-693">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="63fe3-694">Referência da API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="63fe3-694">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
