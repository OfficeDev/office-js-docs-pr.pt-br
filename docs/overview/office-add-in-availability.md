---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos compatíveis com Excel, Word, Outlook, PowerPoint, OneNote e Project.
ms.date: 11/07/2018
ms.openlocfilehash: f8d7d9d393531301829b31dd171a5332a0da536b
ms.sourcegitcommit: 9b021af6cb23a58486d6c5c7492be425e309bea1
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/15/2018
ms.locfileid: "26533795"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="8f33c-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="8f33c-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="8f33c-104">Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="8f33c-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="8f33c-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que atualmente são compatíveis com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="8f33c-105">The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="8f33c-p102">O número do build do Office 2016 instalado via MSI é 16.0.4266.1001. Esta versão só contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="8f33c-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="8f33c-108">Excel</span><span class="sxs-lookup"><span data-stu-id="8f33c-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="8f33c-109">Plataforma</span><span class="sxs-lookup"><span data-stu-id="8f33c-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="8f33c-110">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="8f33c-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="8f33c-111">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="8f33c-111">Identity API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="8f33c-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="8f33c-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="8f33c-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="8f33c-113">Office Online</span></span></td>
    <td> <span data-ttu-id="8f33c-114">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="8f33c-114">-TaskPane</span></span><br><span data-ttu-id="8f33c-115">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="8f33c-115">
        - Content</span></span><br><span data-ttu-id="8f33c-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="8f33c-116">Add-in commands</span></span></td>
    <td><span data-ttu-id="8f33c-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="8f33c-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="8f33c-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="8f33c-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="8f33c-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="8f33c-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="8f33c-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-123">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="8f33c-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="8f33c-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="8f33c-126">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8f33c-126">
        -BindingEvents</span></span><br><span data-ttu-id="8f33c-127">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-127">
        -CompressedFile</span></span><br><span data-ttu-id="8f33c-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8f33c-128">
        -DocumentEvents</span></span><br><span data-ttu-id="8f33c-129">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8f33c-129">
        -</span></span><br><span data-ttu-id="8f33c-130">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-130">
        -MatrixBindings</span></span><br><span data-ttu-id="8f33c-131">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-131">
        -MatrixCoercion</span></span><br><span data-ttu-id="8f33c-132">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="8f33c-132">
        - Selection</span></span><br><span data-ttu-id="8f33c-133">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="8f33c-133">
        - Settings</span></span><br><span data-ttu-id="8f33c-134">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-134">
        -TableBindings</span></span><br><span data-ttu-id="8f33c-135">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-135">
        -TableCoercion</span></span><br><span data-ttu-id="8f33c-136">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-136">
        -TextBindings</span></span><br><span data-ttu-id="8f33c-137">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-137">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8f33c-138">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="8f33c-138">Outlook 2013 for Windows</span></span></td>
    <td><span data-ttu-id="8f33c-139">
        - TaskPane</span><span class="sxs-lookup"><span data-stu-id="8f33c-139">
        -TaskPane</span></span><br><span data-ttu-id="8f33c-140">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="8f33c-140">
        - Content</span></span></td>
    <td>  <span data-ttu-id="8f33c-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="8f33c-142">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8f33c-142">
        -BindingEvents</span></span><br><span data-ttu-id="8f33c-143">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-143">
        -CompressedFile</span></span><br><span data-ttu-id="8f33c-144">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8f33c-144">
        -DocumentEvents</span></span><br><span data-ttu-id="8f33c-145">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8f33c-145">
        -</span></span><br><span data-ttu-id="8f33c-146">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-146">
        -ImageCoercion</span></span><br><span data-ttu-id="8f33c-147">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-147">
        -MatrixBindings</span></span><br><span data-ttu-id="8f33c-148">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-148">
        -MatrixCoercion</span></span><br><span data-ttu-id="8f33c-149">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="8f33c-149">
        - Selection</span></span><br><span data-ttu-id="8f33c-150">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="8f33c-150">
        - Settings</span></span><br><span data-ttu-id="8f33c-151">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-151">
        -TableBindings</span></span><br><span data-ttu-id="8f33c-152">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-152">
        -TableCoercion</span></span><br><span data-ttu-id="8f33c-153">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-153">
        -TextBindings</span></span><br><span data-ttu-id="8f33c-154">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-154">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8f33c-155">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="8f33c-155">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="8f33c-156">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="8f33c-156">-TaskPane</span></span><br><span data-ttu-id="8f33c-157">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="8f33c-157">
        - Content</span></span><br><span data-ttu-id="8f33c-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in commands</a></span></span></td>
    <td><span data-ttu-id="8f33c-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="8f33c-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="8f33c-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="8f33c-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="8f33c-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="8f33c-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="8f33c-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-165">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="8f33c-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="8f33c-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="8f33c-168">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8f33c-168">-BindingEvents</span></span><br><span data-ttu-id="8f33c-169">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-169">
        -CompressedFile</span></span><br><span data-ttu-id="8f33c-170">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8f33c-170">
        -DocumentEvents</span></span><br><span data-ttu-id="8f33c-171">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8f33c-171">
        -</span></span><br><span data-ttu-id="8f33c-172">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-172">
        -ImageCoercion</span></span><br><span data-ttu-id="8f33c-173">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-173">
        -MatrixBindings</span></span><br><span data-ttu-id="8f33c-174">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-174">
        -MatrixCoercion</span></span><br><span data-ttu-id="8f33c-175">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="8f33c-175">
        - Selection</span></span><br><span data-ttu-id="8f33c-176">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="8f33c-176">
        - Settings</span></span><br><span data-ttu-id="8f33c-177">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-177">
        -TableBindings</span></span><br><span data-ttu-id="8f33c-178">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-178">
        -TableCoercion</span></span><br><span data-ttu-id="8f33c-179">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-179">
        -TextBindings</span></span><br><span data-ttu-id="8f33c-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-180">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8f33c-181">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="8f33c-181">Office for Windows Desktop.</span></span></td>
    <td><span data-ttu-id="8f33c-182">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="8f33c-182">-TaskPane</span></span><br><span data-ttu-id="8f33c-183">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="8f33c-183">
        - Content</span></span><br><span data-ttu-id="8f33c-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in commands</a></span></span></td>
    <td><span data-ttu-id="8f33c-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="8f33c-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="8f33c-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="8f33c-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="8f33c-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="8f33c-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="8f33c-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-191">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="8f33c-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="8f33c-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="8f33c-194">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8f33c-194">-BindingEvents</span></span><br><span data-ttu-id="8f33c-195">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-195">
        -CompressedFile</span></span><br><span data-ttu-id="8f33c-196">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8f33c-196">
        -DocumentEvents</span></span><br><span data-ttu-id="8f33c-197">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8f33c-197">
        -</span></span><br><span data-ttu-id="8f33c-198">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-198">
        -ImageCoercion</span></span><br><span data-ttu-id="8f33c-199">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-199">
        -MatrixBindings</span></span><br><span data-ttu-id="8f33c-200">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-200">
        -MatrixCoercion</span></span><br><span data-ttu-id="8f33c-201">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="8f33c-201">
        - Selection</span></span><br><span data-ttu-id="8f33c-202">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="8f33c-202">
        - Settings</span></span><br><span data-ttu-id="8f33c-203">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-203">
        -TableBindings</span></span><br><span data-ttu-id="8f33c-204">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-204">
        -TableCoercion</span></span><br><span data-ttu-id="8f33c-205">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-205">
        -TextBindings</span></span><br><span data-ttu-id="8f33c-206">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-206">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8f33c-207">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="8f33c-207">Office for iOS</span></span></td>
    <td><span data-ttu-id="8f33c-208">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="8f33c-208">-TaskPane</span></span><br><span data-ttu-id="8f33c-209">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="8f33c-209">
        - Content</span></span></td>
    <td><span data-ttu-id="8f33c-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="8f33c-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="8f33c-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="8f33c-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="8f33c-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="8f33c-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="8f33c-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-216">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="8f33c-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="8f33c-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="8f33c-219">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8f33c-219">-BindingEvents</span></span><br><span data-ttu-id="8f33c-220">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-220">
        -CompressedFile</span></span><br><span data-ttu-id="8f33c-221">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8f33c-221">
        -DocumentEvents</span></span><br><span data-ttu-id="8f33c-222">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8f33c-222">
        -</span></span><br><span data-ttu-id="8f33c-223">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-223">
        -ImageCoercion</span></span><br><span data-ttu-id="8f33c-224">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-224">
        -MatrixBindings</span></span><br><span data-ttu-id="8f33c-225">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-225">
        -MatrixCoercion</span></span><br><span data-ttu-id="8f33c-226">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="8f33c-226">
        - Selection</span></span><br><span data-ttu-id="8f33c-227">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="8f33c-227">
        - Settings</span></span><br><span data-ttu-id="8f33c-228">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-228">
        -TableBindings</span></span><br><span data-ttu-id="8f33c-229">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-229">
        -TableCoercion</span></span><br><span data-ttu-id="8f33c-230">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-230">
        -TextBindings</span></span><br><span data-ttu-id="8f33c-231">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-231">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8f33c-232">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="8f33c-232">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="8f33c-233">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="8f33c-233">-TaskPane</span></span><br><span data-ttu-id="8f33c-234">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="8f33c-234">
        - Content</span></span><br><span data-ttu-id="8f33c-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in commands</a></span></span></td>
    <td><span data-ttu-id="8f33c-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="8f33c-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="8f33c-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="8f33c-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="8f33c-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="8f33c-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="8f33c-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-242">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="8f33c-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="8f33c-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="8f33c-245">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8f33c-245">-BindingEvents</span></span><br><span data-ttu-id="8f33c-246">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-246">
        -CompressedFile</span></span><br><span data-ttu-id="8f33c-247">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8f33c-247">
        -DocumentEvents</span></span><br><span data-ttu-id="8f33c-248">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8f33c-248">
        -</span></span><br><span data-ttu-id="8f33c-249">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-249">
        -ImageCoercion</span></span><br><span data-ttu-id="8f33c-250">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-250">
        -MatrixBindings</span></span><br><span data-ttu-id="8f33c-251">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-251">
        -MatrixCoercion</span></span><br><span data-ttu-id="8f33c-252">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-252">
        -PdfFile</span></span><br><span data-ttu-id="8f33c-253">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="8f33c-253">
        - Selection</span></span><br><span data-ttu-id="8f33c-254">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="8f33c-254">
        - Settings</span></span><br><span data-ttu-id="8f33c-255">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-255">
        -TableBindings</span></span><br><span data-ttu-id="8f33c-256">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-256">
        -TableCoercion</span></span><br><span data-ttu-id="8f33c-257">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-257">
        -TextBindings</span></span><br><span data-ttu-id="8f33c-258">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-258">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8f33c-259">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="8f33c-259">Office for Mac</span></span></td>
    <td><span data-ttu-id="8f33c-260">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="8f33c-260">-TaskPane</span></span><br><span data-ttu-id="8f33c-261">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="8f33c-261">
        - Content</span></span><br><span data-ttu-id="8f33c-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in commands</a></span></span></td>
    <td><span data-ttu-id="8f33c-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="8f33c-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="8f33c-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="8f33c-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="8f33c-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="8f33c-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="8f33c-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-269">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="8f33c-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="8f33c-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="8f33c-272">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8f33c-272">-BindingEvents</span></span><br><span data-ttu-id="8f33c-273">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-273">
        -CompressedFile</span></span><br><span data-ttu-id="8f33c-274">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8f33c-274">
        -DocumentEvents</span></span><br><span data-ttu-id="8f33c-275">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8f33c-275">
        -</span></span><br><span data-ttu-id="8f33c-276">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-276">
        -ImageCoercion</span></span><br><span data-ttu-id="8f33c-277">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-277">
        -MatrixBindings</span></span><br><span data-ttu-id="8f33c-278">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-278">
        -MatrixCoercion</span></span><br><span data-ttu-id="8f33c-279">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-279">
        -PdfFile</span></span><br><span data-ttu-id="8f33c-280">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="8f33c-280">
        - Selection</span></span><br><span data-ttu-id="8f33c-281">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="8f33c-281">
        - Settings</span></span><br><span data-ttu-id="8f33c-282">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-282">
        -TableBindings</span></span><br><span data-ttu-id="8f33c-283">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-283">
        -TableCoercion</span></span><br><span data-ttu-id="8f33c-284">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-284">
        -TextBindings</span></span><br><span data-ttu-id="8f33c-285">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-285">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="8f33c-286">Outlook</span><span class="sxs-lookup"><span data-stu-id="8f33c-286">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="8f33c-287">Plataforma</span><span class="sxs-lookup"><span data-stu-id="8f33c-287">Platform</span></span></th>
    <th><span data-ttu-id="8f33c-288">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="8f33c-288">Extension points</span></span></th>
    <th><span data-ttu-id="8f33c-289">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="8f33c-289">Identity API requirement sets</span></span></th>
    <th><span data-ttu-id="8f33c-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="8f33c-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="8f33c-291">Office Online</span><span class="sxs-lookup"><span data-stu-id="8f33c-291">Office Online</span></span></td>
    <td> <span data-ttu-id="8f33c-292">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="8f33c-292">- Mail Read</span></span><br><span data-ttu-id="8f33c-293">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="8f33c-293">
      - Mail Compose</span></span><br><span data-ttu-id="8f33c-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in commands</a></span></span></td>
    <td> <span data-ttu-id="8f33c-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="8f33c-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"></a> Mailbox 1.2</span></span><br><span data-ttu-id="8f33c-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="8f33c-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="8f33c-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="8f33c-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="8f33c-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="8f33c-302">Não disponível</span><span class="sxs-lookup"><span data-stu-id="8f33c-302">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8f33c-303">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="8f33c-303">Outlook 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="8f33c-304">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="8f33c-304">- Mail Read</span></span><br><span data-ttu-id="8f33c-305">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="8f33c-305">
      - Mail Compose</span></span><br><span data-ttu-id="8f33c-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in commands</a></span></span></td>
    <td> <span data-ttu-id="8f33c-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="8f33c-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"></a> Mailbox 1.2</span></span><br><span data-ttu-id="8f33c-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="8f33c-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="8f33c-311">Não disponível</span><span class="sxs-lookup"><span data-stu-id="8f33c-311">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8f33c-312">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="8f33c-312">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="8f33c-313">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="8f33c-313">- Mail Read</span></span><br><span data-ttu-id="8f33c-314">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="8f33c-314">
      - Mail Compose</span></span><br><span data-ttu-id="8f33c-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in commands</a></span></span><br><span data-ttu-id="8f33c-316">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="8f33c-316">
      - Modules</span></span></td>
    <td> <span data-ttu-id="8f33c-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="8f33c-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"></a> Mailbox 1.2</span></span><br><span data-ttu-id="8f33c-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="8f33c-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="8f33c-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="8f33c-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="8f33c-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="8f33c-324">Não disponível</span><span class="sxs-lookup"><span data-stu-id="8f33c-324">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8f33c-325">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="8f33c-325">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="8f33c-326">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="8f33c-326">- Mail Read</span></span><br><span data-ttu-id="8f33c-327">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="8f33c-327">
      - Mail Compose</span></span><br><span data-ttu-id="8f33c-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in commands</a></span></span><br><span data-ttu-id="8f33c-329">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="8f33c-329">
      - Modules</span></span></td>
    <td> <span data-ttu-id="8f33c-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="8f33c-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"></a> Mailbox 1.2</span></span><br><span data-ttu-id="8f33c-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="8f33c-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="8f33c-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="8f33c-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="8f33c-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="8f33c-337">Não disponível</span><span class="sxs-lookup"><span data-stu-id="8f33c-337">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8f33c-338">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="8f33c-338">Office for iOS</span></span></td>
    <td> <span data-ttu-id="8f33c-339">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="8f33c-339">- Mail Read</span></span><br><span data-ttu-id="8f33c-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in commands</a></span></span></td>
    <td> <span data-ttu-id="8f33c-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="8f33c-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"></a> Mailbox 1.2</span></span><br><span data-ttu-id="8f33c-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="8f33c-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="8f33c-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="8f33c-346">Não disponível</span><span class="sxs-lookup"><span data-stu-id="8f33c-346">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8f33c-347">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="8f33c-347">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="8f33c-348">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="8f33c-348">- Mail Read</span></span><br><span data-ttu-id="8f33c-349">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="8f33c-349">
      - Mail Compose</span></span><br><span data-ttu-id="8f33c-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in commands</a></span></span></td>
    <td> <span data-ttu-id="8f33c-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="8f33c-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"></a> Mailbox 1.2</span></span><br><span data-ttu-id="8f33c-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="8f33c-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="8f33c-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="8f33c-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="8f33c-357">Não disponível</span><span class="sxs-lookup"><span data-stu-id="8f33c-357">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8f33c-358">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="8f33c-358">Office for Mac</span></span></td>
    <td> <span data-ttu-id="8f33c-359">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="8f33c-359">- Mail Read</span></span><br><span data-ttu-id="8f33c-360">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="8f33c-360">
      - Mail Compose</span></span><br><span data-ttu-id="8f33c-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in commands</a></span></span></td>
    <td> <span data-ttu-id="8f33c-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="8f33c-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"></a> Mailbox 1.2</span></span><br><span data-ttu-id="8f33c-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="8f33c-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="8f33c-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="8f33c-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="8f33c-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="8f33c-369">Não disponível</span><span class="sxs-lookup"><span data-stu-id="8f33c-369">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8f33c-370">Office para Android</span><span class="sxs-lookup"><span data-stu-id="8f33c-370">Office for Android</span></span></td>
    <td> <span data-ttu-id="8f33c-371">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="8f33c-371">- Mail Read</span></span><br><span data-ttu-id="8f33c-372">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-372">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in commands</a></span></span></td>
    <td> <span data-ttu-id="8f33c-373">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-373">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="8f33c-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"></a> Mailbox 1.2</span></span><br><span data-ttu-id="8f33c-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="8f33c-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="8f33c-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="8f33c-378">Não disponível</span><span class="sxs-lookup"><span data-stu-id="8f33c-378">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="8f33c-379">Word</span><span class="sxs-lookup"><span data-stu-id="8f33c-379">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="8f33c-380">Plataforma</span><span class="sxs-lookup"><span data-stu-id="8f33c-380">Platform</span></span></th>
    <th><span data-ttu-id="8f33c-381">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="8f33c-381">Extension points</span></span></th>
    <th><span data-ttu-id="8f33c-382">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="8f33c-382">Identity API requirement sets</span></span></th>
    <th><span data-ttu-id="8f33c-383"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="8f33c-383"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="8f33c-384">Office Online</span><span class="sxs-lookup"><span data-stu-id="8f33c-384">Office Online</span></span></td>
    <td> <span data-ttu-id="8f33c-385">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="8f33c-385">-TaskPane</span></span><br><span data-ttu-id="8f33c-386">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-386">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in commands</a></span></span></td>
    <td> <span data-ttu-id="8f33c-387">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-387">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="8f33c-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="8f33c-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="8f33c-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="8f33c-391">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8f33c-391">-BindingEvents</span></span><br><span data-ttu-id="8f33c-392">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="8f33c-392">
         -</span></span><br><span data-ttu-id="8f33c-393">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8f33c-393">
         -DocumentEvents</span></span><br><span data-ttu-id="8f33c-394">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8f33c-394">
         -</span></span><br><span data-ttu-id="8f33c-395">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-395">
         -HtmlCoercion</span></span><br><span data-ttu-id="8f33c-396">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-396">
         -ImageCoercion</span></span><br><span data-ttu-id="8f33c-397">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-397">
         -MatrixBindings</span></span><br><span data-ttu-id="8f33c-398">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-398">
         -MatrixCoercion</span></span><br><span data-ttu-id="8f33c-399">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-399">
         -OoxmlCoercion</span></span><br><span data-ttu-id="8f33c-400">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-400">
         -PdfFile</span></span><br><span data-ttu-id="8f33c-401">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="8f33c-401">
         - Selection</span></span><br><span data-ttu-id="8f33c-402">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="8f33c-402">
         - Settings</span></span><br><span data-ttu-id="8f33c-403">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-403">
         -TableBindings</span></span><br><span data-ttu-id="8f33c-404">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-404">
         -TableCoercion</span></span><br><span data-ttu-id="8f33c-405">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-405">
         -TextBindings</span></span><br><span data-ttu-id="8f33c-406">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-406">
         -TextCoercion</span></span><br><span data-ttu-id="8f33c-407">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-407">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8f33c-408">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="8f33c-408">Outlook 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="8f33c-409">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="8f33c-409">-TaskPane</span></span></td>
    <td> <span data-ttu-id="8f33c-410">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-410">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="8f33c-411">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8f33c-411">-BindingEvents</span></span><br><span data-ttu-id="8f33c-412">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-412">
         -CompressedFile</span></span><br><span data-ttu-id="8f33c-413">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="8f33c-413">
         -</span></span><br><span data-ttu-id="8f33c-414">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8f33c-414">
         -DocumentEvents</span></span><br><span data-ttu-id="8f33c-415">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8f33c-415">
         -</span></span><br><span data-ttu-id="8f33c-416">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-416">
         -HtmlCoercion</span></span><br><span data-ttu-id="8f33c-417">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-417">
         -ImageCoercion</span></span><br><span data-ttu-id="8f33c-418">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-418">
         -MatrixBindings</span></span><br><span data-ttu-id="8f33c-419">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-419">
         -MatrixCoercion</span></span><br><span data-ttu-id="8f33c-420">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-420">
         -OoxmlCoercion</span></span><br><span data-ttu-id="8f33c-421">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-421">
         -PdfFile</span></span><br><span data-ttu-id="8f33c-422">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="8f33c-422">
         - Selection</span></span><br><span data-ttu-id="8f33c-423">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="8f33c-423">
         - Settings</span></span><br><span data-ttu-id="8f33c-424">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-424">
         -TableBindings</span></span><br><span data-ttu-id="8f33c-425">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-425">
         -TableCoercion</span></span><br><span data-ttu-id="8f33c-426">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-426">
         -TextBindings</span></span><br><span data-ttu-id="8f33c-427">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-427">
         -TextCoercion</span></span><br><span data-ttu-id="8f33c-428">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-428">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8f33c-429">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="8f33c-429">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="8f33c-430">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="8f33c-430">-TaskPane</span></span><br><span data-ttu-id="8f33c-431">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-431">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in commands</a></span></span></td>
    <td> <span data-ttu-id="8f33c-432">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-432">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="8f33c-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="8f33c-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="8f33c-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="8f33c-436">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8f33c-436">-BindingEvents</span></span><br><span data-ttu-id="8f33c-437">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-437">
         -CompressedFile</span></span><br><span data-ttu-id="8f33c-438">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="8f33c-438">
         -</span></span><br><span data-ttu-id="8f33c-439">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8f33c-439">
         -DocumentEvents</span></span><br><span data-ttu-id="8f33c-440">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8f33c-440">
         -</span></span><br><span data-ttu-id="8f33c-441">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-441">
         -HtmlCoercion</span></span><br><span data-ttu-id="8f33c-442">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-442">
         -ImageCoercion</span></span><br><span data-ttu-id="8f33c-443">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-443">
         -MatrixBindings</span></span><br><span data-ttu-id="8f33c-444">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-444">
         -MatrixCoercion</span></span><br><span data-ttu-id="8f33c-445">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-445">
         -OoxmlCoercion</span></span><br><span data-ttu-id="8f33c-446">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-446">
         -PdfFile</span></span><br><span data-ttu-id="8f33c-447">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="8f33c-447">
         - Selection</span></span><br><span data-ttu-id="8f33c-448">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="8f33c-448">
         - Settings</span></span><br><span data-ttu-id="8f33c-449">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-449">
         -TableBindings</span></span><br><span data-ttu-id="8f33c-450">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-450">
         -TableCoercion</span></span><br><span data-ttu-id="8f33c-451">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-451">
         -TextBindings</span></span><br><span data-ttu-id="8f33c-452">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-452">
         -TextCoercion</span></span><br><span data-ttu-id="8f33c-453">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-453">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="8f33c-454">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="8f33c-454">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="8f33c-455">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="8f33c-455">-TaskPane</span></span><br><span data-ttu-id="8f33c-456">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-456">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in commands</a></span></span></td>
    <td> <span data-ttu-id="8f33c-457">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-457">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="8f33c-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="8f33c-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="8f33c-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="8f33c-461">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8f33c-461">-BindingEvents</span></span><br><span data-ttu-id="8f33c-462">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-462">
         -CompressedFile</span></span><br><span data-ttu-id="8f33c-463">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="8f33c-463">
         -</span></span><br><span data-ttu-id="8f33c-464">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8f33c-464">
         -DocumentEvents</span></span><br><span data-ttu-id="8f33c-465">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8f33c-465">
         -</span></span><br><span data-ttu-id="8f33c-466">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-466">
         -HtmlCoercion</span></span><br><span data-ttu-id="8f33c-467">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-467">
         -ImageCoercion</span></span><br><span data-ttu-id="8f33c-468">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-468">
         -MatrixBindings</span></span><br><span data-ttu-id="8f33c-469">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-469">
         -MatrixCoercion</span></span><br><span data-ttu-id="8f33c-470">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-470">
         -OoxmlCoercion</span></span><br><span data-ttu-id="8f33c-471">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-471">
         -PdfFile</span></span><br><span data-ttu-id="8f33c-472">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="8f33c-472">
         - Selection</span></span><br><span data-ttu-id="8f33c-473">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="8f33c-473">
         - Settings</span></span><br><span data-ttu-id="8f33c-474">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-474">
         -TableBindings</span></span><br><span data-ttu-id="8f33c-475">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-475">
         -TableCoercion</span></span><br><span data-ttu-id="8f33c-476">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-476">
         -TextBindings</span></span><br><span data-ttu-id="8f33c-477">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-477">
         -TextCoercion</span></span><br><span data-ttu-id="8f33c-478">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-478">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="8f33c-479">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="8f33c-479">Office for iOS</span></span></td>
    <td> <span data-ttu-id="8f33c-480">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="8f33c-480">-TaskPane</span></span></td>
    <td> <span data-ttu-id="8f33c-481">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-481">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="8f33c-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="8f33c-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="8f33c-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="8f33c-484">DialogApi 1.1</span></span></td>
    <td> <span data-ttu-id="8f33c-485">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8f33c-485">-BindingEvents</span></span><br><span data-ttu-id="8f33c-486">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-486">
         -CompressedFile</span></span><br><span data-ttu-id="8f33c-487">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="8f33c-487">
         -</span></span><br><span data-ttu-id="8f33c-488">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8f33c-488">
         -DocumentEvents</span></span><br><span data-ttu-id="8f33c-489">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8f33c-489">
         -</span></span><br><span data-ttu-id="8f33c-490">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-490">
         -HtmlCoercion</span></span><br><span data-ttu-id="8f33c-491">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-491">
         -ImageCoercion</span></span><br><span data-ttu-id="8f33c-492">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-492">
         -MatrixBindings</span></span><br><span data-ttu-id="8f33c-493">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-493">
         -MatrixCoercion</span></span><br><span data-ttu-id="8f33c-494">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-494">
         -OoxmlCoercion</span></span><br><span data-ttu-id="8f33c-495">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-495">
         -PdfFile</span></span><br><span data-ttu-id="8f33c-496">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="8f33c-496">
         - Selection</span></span><br><span data-ttu-id="8f33c-497">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="8f33c-497">
         - Settings</span></span><br><span data-ttu-id="8f33c-498">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-498">
         -TableBindings</span></span><br><span data-ttu-id="8f33c-499">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-499">
         -TableCoercion</span></span><br><span data-ttu-id="8f33c-500">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-500">
         -TextBindings</span></span><br><span data-ttu-id="8f33c-501">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-501">
         -TextCoercion</span></span><br><span data-ttu-id="8f33c-502">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-502">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="8f33c-503">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="8f33c-503">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="8f33c-504">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="8f33c-504">-TaskPane</span></span><br><span data-ttu-id="8f33c-505">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-505">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in commands</a></span></span></td>
    <td> <span data-ttu-id="8f33c-506">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-506">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="8f33c-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="8f33c-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="8f33c-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="8f33c-509">DialogApi 1.1</span></span></td>
    <td> <span data-ttu-id="8f33c-510">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8f33c-510">-BindingEvents</span></span><br><span data-ttu-id="8f33c-511">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-511">
         -CompressedFile</span></span><br><span data-ttu-id="8f33c-512">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="8f33c-512">
         -</span></span><br><span data-ttu-id="8f33c-513">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8f33c-513">
         -DocumentEvents</span></span><br><span data-ttu-id="8f33c-514">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8f33c-514">
         -</span></span><br><span data-ttu-id="8f33c-515">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-515">
         -HtmlCoercion</span></span><br><span data-ttu-id="8f33c-516">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-516">
         -ImageCoercion</span></span><br><span data-ttu-id="8f33c-517">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-517">
         -MatrixBindings</span></span><br><span data-ttu-id="8f33c-518">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-518">
         -MatrixCoercion</span></span><br><span data-ttu-id="8f33c-519">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-519">
         -OoxmlCoercion</span></span><br><span data-ttu-id="8f33c-520">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-520">
         -PdfFile</span></span><br><span data-ttu-id="8f33c-521">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="8f33c-521">
         - Selection</span></span><br><span data-ttu-id="8f33c-522">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="8f33c-522">
         - Settings</span></span><br><span data-ttu-id="8f33c-523">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-523">
         -TableBindings</span></span><br><span data-ttu-id="8f33c-524">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-524">
         -TableCoercion</span></span><br><span data-ttu-id="8f33c-525">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-525">
         -TextBindings</span></span><br><span data-ttu-id="8f33c-526">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-526">
         -TextCoercion</span></span><br><span data-ttu-id="8f33c-527">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-527">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="8f33c-528">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="8f33c-528">Office for Mac</span></span></td>
    <td> <span data-ttu-id="8f33c-529">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="8f33c-529">-TaskPane</span></span><br><span data-ttu-id="8f33c-530">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-530">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in commands</a></span></span></td>
    <td> <span data-ttu-id="8f33c-531">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-531">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="8f33c-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="8f33c-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="8f33c-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="8f33c-534">DialogApi 1.1</span></span></td>
    <td> <span data-ttu-id="8f33c-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8f33c-535">-BindingEvents</span></span><br><span data-ttu-id="8f33c-536">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-536">
         -CompressedFile</span></span><br><span data-ttu-id="8f33c-537">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="8f33c-537">
         -</span></span><br><span data-ttu-id="8f33c-538">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8f33c-538">
         -DocumentEvents</span></span><br><span data-ttu-id="8f33c-539">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8f33c-539">
         -</span></span><br><span data-ttu-id="8f33c-540">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-540">
         -HtmlCoercion</span></span><br><span data-ttu-id="8f33c-541">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-541">
         -ImageCoercion</span></span><br><span data-ttu-id="8f33c-542">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-542">
         -MatrixBindings</span></span><br><span data-ttu-id="8f33c-543">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-543">
         -MatrixCoercion</span></span><br><span data-ttu-id="8f33c-544">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-544">
         -OoxmlCoercion</span></span><br><span data-ttu-id="8f33c-545">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-545">
         -PdfFile</span></span><br><span data-ttu-id="8f33c-546">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="8f33c-546">
         - Selection</span></span><br><span data-ttu-id="8f33c-547">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="8f33c-547">
         - Settings</span></span><br><span data-ttu-id="8f33c-548">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-548">
         -TableBindings</span></span><br><span data-ttu-id="8f33c-549">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-549">
         -TableCoercion</span></span><br><span data-ttu-id="8f33c-550">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8f33c-550">
         -TextBindings</span></span><br><span data-ttu-id="8f33c-551">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-551">
         -TextCoercion</span></span><br><span data-ttu-id="8f33c-552">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-552">
         -TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="8f33c-553">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="8f33c-553">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="8f33c-554">Plataforma</span><span class="sxs-lookup"><span data-stu-id="8f33c-554">Platform</span></span></th>
    <th><span data-ttu-id="8f33c-555">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="8f33c-555">Extension points</span></span></th>
    <th><span data-ttu-id="8f33c-556">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="8f33c-556">Identity API requirement sets</span></span></th>
    <th><span data-ttu-id="8f33c-557"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="8f33c-557"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="8f33c-558">Office Online</span><span class="sxs-lookup"><span data-stu-id="8f33c-558">Office Online</span></span></td>
    <td> <span data-ttu-id="8f33c-559">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="8f33c-559">- Content</span></span><br><span data-ttu-id="8f33c-560">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="8f33c-560">
         -TaskPane</span></span><br><span data-ttu-id="8f33c-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in commands</a></span></span></td>
    <td> <span data-ttu-id="8f33c-562">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-562">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="8f33c-563">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="8f33c-563">-</span></span><br><span data-ttu-id="8f33c-564">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-564">
         -CompressedFile</span></span><br><span data-ttu-id="8f33c-565">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8f33c-565">
         -DocumentEvents</span></span><br><span data-ttu-id="8f33c-566">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8f33c-566">
         -</span></span><br><span data-ttu-id="8f33c-567">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-567">
         -ImageCoercion</span></span><br><span data-ttu-id="8f33c-568">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-568">
         -PdfFile</span></span><br><span data-ttu-id="8f33c-569">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="8f33c-569">
         - Selection</span></span><br><span data-ttu-id="8f33c-570">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="8f33c-570">
         - Settings</span></span><br><span data-ttu-id="8f33c-571">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-571">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8f33c-572">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="8f33c-572">Outlook 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="8f33c-573">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="8f33c-573">- Content</span></span><br><span data-ttu-id="8f33c-574">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="8f33c-574">
         -TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="8f33c-575">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="8f33c-575">DialogApi 1.1</span></span></td>
    <td> <span data-ttu-id="8f33c-576">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="8f33c-576">-</span></span><br><span data-ttu-id="8f33c-577">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-577">
         -CompressedFile</span></span><br><span data-ttu-id="8f33c-578">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8f33c-578">
         -DocumentEvents</span></span><br><span data-ttu-id="8f33c-579">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8f33c-579">
         -</span></span><br><span data-ttu-id="8f33c-580">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-580">
         -ImageCoercion</span></span><br><span data-ttu-id="8f33c-581">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-581">
         -PdfFile</span></span><br><span data-ttu-id="8f33c-582">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="8f33c-582">
         - Selection</span></span><br><span data-ttu-id="8f33c-583">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="8f33c-583">
         - Settings</span></span><br><span data-ttu-id="8f33c-584">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-584">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8f33c-585">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="8f33c-585">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="8f33c-586">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="8f33c-586">- Content</span></span><br><span data-ttu-id="8f33c-587">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="8f33c-587">
         -TaskPane</span></span><br><span data-ttu-id="8f33c-588">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-588">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in commands</a></span></span></td>
    <td> <span data-ttu-id="8f33c-589">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-589">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="8f33c-590">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="8f33c-590">-</span></span><br><span data-ttu-id="8f33c-591">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-591">
         -CompressedFile</span></span><br><span data-ttu-id="8f33c-592">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8f33c-592">
         -DocumentEvents</span></span><br><span data-ttu-id="8f33c-593">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8f33c-593">
         -</span></span><br><span data-ttu-id="8f33c-594">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-594">
         -ImageCoercion</span></span><br><span data-ttu-id="8f33c-595">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-595">
         -PdfFile</span></span><br><span data-ttu-id="8f33c-596">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="8f33c-596">
         - Selection</span></span><br><span data-ttu-id="8f33c-597">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="8f33c-597">
         - Settings</span></span><br><span data-ttu-id="8f33c-598">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-598">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8f33c-599">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="8f33c-599">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="8f33c-600">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="8f33c-600">- Content</span></span><br><span data-ttu-id="8f33c-601">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="8f33c-601">
         -TaskPane</span></span><br><span data-ttu-id="8f33c-602">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-602">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in commands</a></span></span></td>
    <td> <span data-ttu-id="8f33c-603">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-603">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="8f33c-604">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="8f33c-604">-</span></span><br><span data-ttu-id="8f33c-605">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-605">
         -CompressedFile</span></span><br><span data-ttu-id="8f33c-606">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8f33c-606">
         -DocumentEvents</span></span><br><span data-ttu-id="8f33c-607">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8f33c-607">
         -</span></span><br><span data-ttu-id="8f33c-608">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-608">
         -ImageCoercion</span></span><br><span data-ttu-id="8f33c-609">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-609">
         -PdfFile</span></span><br><span data-ttu-id="8f33c-610">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="8f33c-610">
         - Selection</span></span><br><span data-ttu-id="8f33c-611">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="8f33c-611">
         - Settings</span></span><br><span data-ttu-id="8f33c-612">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-612">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8f33c-613">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="8f33c-613">Office for iOS</span></span></td>
    <td> <span data-ttu-id="8f33c-614">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="8f33c-614">- Content</span></span><br><span data-ttu-id="8f33c-615">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="8f33c-615">
         -TaskPane</span></span></td>
    <td> <span data-ttu-id="8f33c-616">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-616">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="8f33c-617">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="8f33c-617">-</span></span><br><span data-ttu-id="8f33c-618">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-618">
         -CompressedFile</span></span><br><span data-ttu-id="8f33c-619">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8f33c-619">
         -DocumentEvents</span></span><br><span data-ttu-id="8f33c-620">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8f33c-620">
         -</span></span><br><span data-ttu-id="8f33c-621">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-621">
         -PdfFile</span></span><br><span data-ttu-id="8f33c-622">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="8f33c-622">
         - Selection</span></span><br><span data-ttu-id="8f33c-623">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="8f33c-623">
         - Settings</span></span><br><span data-ttu-id="8f33c-624">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-624">
         -TextCoercion</span></span><br><span data-ttu-id="8f33c-625">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-625">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8f33c-626">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="8f33c-626">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="8f33c-627">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="8f33c-627">- Content</span></span><br><span data-ttu-id="8f33c-628">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="8f33c-628">
         -TaskPane</span></span><br><span data-ttu-id="8f33c-629">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-629">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in commands</a></span></span></td>
    <td> <span data-ttu-id="8f33c-630">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-630">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="8f33c-631">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="8f33c-631">-</span></span><br><span data-ttu-id="8f33c-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-632">
         -CompressedFile</span></span><br><span data-ttu-id="8f33c-633">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8f33c-633">
         -DocumentEvents</span></span><br><span data-ttu-id="8f33c-634">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8f33c-634">
         -</span></span><br><span data-ttu-id="8f33c-635">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-635">
         -ImageCoercion</span></span><br><span data-ttu-id="8f33c-636">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-636">
         -PdfFile</span></span><br><span data-ttu-id="8f33c-637">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="8f33c-637">
         - Selection</span></span><br><span data-ttu-id="8f33c-638">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="8f33c-638">
         - Settings</span></span><br><span data-ttu-id="8f33c-639">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-639">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8f33c-640">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="8f33c-640">Office for Mac</span></span></td>
    <td> <span data-ttu-id="8f33c-641">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="8f33c-641">- Content</span></span><br><span data-ttu-id="8f33c-642">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="8f33c-642">
         -TaskPane</span></span><br><span data-ttu-id="8f33c-643">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-643">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in commands</a></span></span></td>
    <td> <span data-ttu-id="8f33c-644">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-644">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="8f33c-645">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="8f33c-645">-</span></span><br><span data-ttu-id="8f33c-646">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-646">
         -CompressedFile</span></span><br><span data-ttu-id="8f33c-647">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8f33c-647">
         -DocumentEvents</span></span><br><span data-ttu-id="8f33c-648">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8f33c-648">
         -</span></span><br><span data-ttu-id="8f33c-649">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-649">
         -ImageCoercion</span></span><br><span data-ttu-id="8f33c-650">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8f33c-650">
         -PdfFile</span></span><br><span data-ttu-id="8f33c-651">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="8f33c-651">
         - Selection</span></span><br><span data-ttu-id="8f33c-652">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="8f33c-652">
         - Settings</span></span><br><span data-ttu-id="8f33c-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-653">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="8f33c-654">OneNote</span><span class="sxs-lookup"><span data-stu-id="8f33c-654">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="8f33c-655">Plataforma</span><span class="sxs-lookup"><span data-stu-id="8f33c-655">Platform</span></span></th>
    <th><span data-ttu-id="8f33c-656">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="8f33c-656">Extension points</span></span></th>
    <th><span data-ttu-id="8f33c-657">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="8f33c-657">Identity API requirement sets</span></span></th>
    <th><span data-ttu-id="8f33c-658"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="8f33c-658"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="8f33c-659">Office Online</span><span class="sxs-lookup"><span data-stu-id="8f33c-659">Office Online</span></span></td>
    <td> <span data-ttu-id="8f33c-660">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="8f33c-660">- Content</span></span><br><span data-ttu-id="8f33c-661">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="8f33c-661">
         -TaskPane</span></span><br><span data-ttu-id="8f33c-662">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-662">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in commands</a></span></span></td>
    <td> <span data-ttu-id="8f33c-663">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-663">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="8f33c-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="8f33c-665">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8f33c-665">-DocumentEvents</span></span><br><span data-ttu-id="8f33c-666">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-666">
         -HtmlCoercion</span></span><br><span data-ttu-id="8f33c-667">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-667">
         -ImageCoercion</span></span><br><span data-ttu-id="8f33c-668">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="8f33c-668">
         - Settings</span></span><br><span data-ttu-id="8f33c-669">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-669">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="8f33c-670">Project</span><span class="sxs-lookup"><span data-stu-id="8f33c-670">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="8f33c-671">Plataforma</span><span class="sxs-lookup"><span data-stu-id="8f33c-671">Platform</span></span></th>
    <th><span data-ttu-id="8f33c-672">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="8f33c-672">Extension points</span></span></th>
    <th><span data-ttu-id="8f33c-673">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="8f33c-673">Identity API requirement sets</span></span></th>
    <th><span data-ttu-id="8f33c-674"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="8f33c-674"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="8f33c-675">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="8f33c-675">Outlook 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="8f33c-676">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="8f33c-676">-TaskPane</span></span></td>
    <td> <span data-ttu-id="8f33c-677">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-677">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="8f33c-678">- Seleção</span><span class="sxs-lookup"><span data-stu-id="8f33c-678">- Selection</span></span><br><span data-ttu-id="8f33c-679">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-679">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8f33c-680">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="8f33c-680">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="8f33c-681">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="8f33c-681">-TaskPane</span></span></td>
    <td> <span data-ttu-id="8f33c-682">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-682">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="8f33c-683">- Seleção</span><span class="sxs-lookup"><span data-stu-id="8f33c-683">- Selection</span></span><br><span data-ttu-id="8f33c-684">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-684">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8f33c-685">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="8f33c-685">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="8f33c-686">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="8f33c-686">-TaskPane</span></span></td>
    <td> <span data-ttu-id="8f33c-687">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8f33c-687">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="8f33c-688">- Seleção</span><span class="sxs-lookup"><span data-stu-id="8f33c-688">- Selection</span></span><br><span data-ttu-id="8f33c-689">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8f33c-689">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="8f33c-690">Confira também</span><span class="sxs-lookup"><span data-stu-id="8f33c-690">See also</span></span>

- [<span data-ttu-id="8f33c-691">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="8f33c-691">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="8f33c-692">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="8f33c-692">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="8f33c-693">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="8f33c-693">Add-in commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="8f33c-694">Referência da API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="8f33c-694">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
