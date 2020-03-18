---
title: Elemento OfficeApp no arquivo de manifesto
description: O elemento OfficeApp é o elemento raiz de um manifesto de suplemento do Office.
ms.date: 02/04/2020
localization_priority: Normal
ms.openlocfilehash: 038933f2d06ee5f485dbdb7dd7abdbd95fb97c7d
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720593"
---
# <a name="officeapp-element"></a><span data-ttu-id="4dbdc-103">Elemento OfficeApp</span><span class="sxs-lookup"><span data-stu-id="4dbdc-103">OfficeApp element</span></span>

<span data-ttu-id="4dbdc-104">O elemento raiz no manifesto de um Suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="4dbdc-104">The root element in the manifest of an Office Add-in.</span></span>

<span data-ttu-id="4dbdc-105">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="4dbdc-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="4dbdc-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="4dbdc-106">Syntax</span></span>

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="4dbdc-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="4dbdc-107">Contained in</span></span>

 <span data-ttu-id="4dbdc-108">_none_</span><span class="sxs-lookup"><span data-stu-id="4dbdc-108">_none_</span></span>

## <a name="must-contain"></a><span data-ttu-id="4dbdc-109">Deve conter</span><span class="sxs-lookup"><span data-stu-id="4dbdc-109">Must contain</span></span>

|<span data-ttu-id="4dbdc-110">**Element**</span><span class="sxs-lookup"><span data-stu-id="4dbdc-110">**Element**</span></span>|<span data-ttu-id="4dbdc-111">**Content**</span><span class="sxs-lookup"><span data-stu-id="4dbdc-111">**Content**</span></span>|<span data-ttu-id="4dbdc-112">**Email**</span><span class="sxs-lookup"><span data-stu-id="4dbdc-112">**Mail**</span></span>|<span data-ttu-id="4dbdc-113">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="4dbdc-113">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="4dbdc-114">Id</span><span class="sxs-lookup"><span data-stu-id="4dbdc-114">Id</span></span>](id.md)|<span data-ttu-id="4dbdc-115">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-115">x</span></span>|<span data-ttu-id="4dbdc-116">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-116">x</span></span>|<span data-ttu-id="4dbdc-117">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-117">x</span></span>|
|[<span data-ttu-id="4dbdc-118">Version</span><span class="sxs-lookup"><span data-stu-id="4dbdc-118">Version</span></span>](version.md)|<span data-ttu-id="4dbdc-119">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-119">x</span></span>|<span data-ttu-id="4dbdc-120">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-120">x</span></span>|<span data-ttu-id="4dbdc-121">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-121">x</span></span>|
|[<span data-ttu-id="4dbdc-122">ProviderName</span><span class="sxs-lookup"><span data-stu-id="4dbdc-122">ProviderName</span></span>](providername.md)|<span data-ttu-id="4dbdc-123">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-123">x</span></span>|<span data-ttu-id="4dbdc-124">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-124">x</span></span>|<span data-ttu-id="4dbdc-125">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-125">x</span></span>|
|[<span data-ttu-id="4dbdc-126">DefaultLocale</span><span class="sxs-lookup"><span data-stu-id="4dbdc-126">DefaultLocale</span></span>](defaultlocale.md)|<span data-ttu-id="4dbdc-127">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-127">x</span></span>|<span data-ttu-id="4dbdc-128">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-128">x</span></span>|<span data-ttu-id="4dbdc-129">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-129">x</span></span>|
|[<span data-ttu-id="4dbdc-130">DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="4dbdc-130">DefaultSettings</span></span>](defaultsettings.md)|<span data-ttu-id="4dbdc-131">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-131">x</span></span>||<span data-ttu-id="4dbdc-132">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-132">x</span></span>|
|[<span data-ttu-id="4dbdc-133">DisplayName</span><span class="sxs-lookup"><span data-stu-id="4dbdc-133">DisplayName</span></span>](displayname.md)|<span data-ttu-id="4dbdc-134">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-134">x</span></span>|<span data-ttu-id="4dbdc-135">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-135">x</span></span>|<span data-ttu-id="4dbdc-136">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-136">x</span></span>|
|[<span data-ttu-id="4dbdc-137">Descrição</span><span class="sxs-lookup"><span data-stu-id="4dbdc-137">Description</span></span>](description.md)|<span data-ttu-id="4dbdc-138">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-138">x</span></span>|<span data-ttu-id="4dbdc-139">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-139">x</span></span>|<span data-ttu-id="4dbdc-140">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-140">x</span></span>|
|[<span data-ttu-id="4dbdc-141">FormSettings</span><span class="sxs-lookup"><span data-stu-id="4dbdc-141">FormSettings</span></span>](formsettings.md)||<span data-ttu-id="4dbdc-142">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-142">x</span></span>||
|[<span data-ttu-id="4dbdc-143">Permissões</span><span class="sxs-lookup"><span data-stu-id="4dbdc-143">Permissions</span></span>](permissions.md)|<span data-ttu-id="4dbdc-144">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-144">x</span></span>||<span data-ttu-id="4dbdc-145">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-145">x</span></span>|
|[<span data-ttu-id="4dbdc-146">Rule</span><span class="sxs-lookup"><span data-stu-id="4dbdc-146">Rule</span></span>](rule.md)||<span data-ttu-id="4dbdc-147">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-147">x</span></span>||

## <a name="can-contain"></a><span data-ttu-id="4dbdc-148">Pode conter</span><span class="sxs-lookup"><span data-stu-id="4dbdc-148">Can contain</span></span>

|<span data-ttu-id="4dbdc-149">**Elemento**</span><span class="sxs-lookup"><span data-stu-id="4dbdc-149">**Element**</span></span>|<span data-ttu-id="4dbdc-150">**Content**</span><span class="sxs-lookup"><span data-stu-id="4dbdc-150">**Content**</span></span>|<span data-ttu-id="4dbdc-151">**Email**</span><span class="sxs-lookup"><span data-stu-id="4dbdc-151">**Mail**</span></span>|<span data-ttu-id="4dbdc-152">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="4dbdc-152">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="4dbdc-153">AlternateId</span><span class="sxs-lookup"><span data-stu-id="4dbdc-153">AlternateId</span></span>](alternateid.md)|<span data-ttu-id="4dbdc-154">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-154">x</span></span>|<span data-ttu-id="4dbdc-155">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-155">x</span></span>|<span data-ttu-id="4dbdc-156">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-156">x</span></span>|
|[<span data-ttu-id="4dbdc-157">IconUrl</span><span class="sxs-lookup"><span data-stu-id="4dbdc-157">IconUrl</span></span>](iconurl.md)|<span data-ttu-id="4dbdc-158">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-158">x</span></span>|<span data-ttu-id="4dbdc-159">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-159">x</span></span>|<span data-ttu-id="4dbdc-160">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-160">x</span></span>|
|[<span data-ttu-id="4dbdc-161">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="4dbdc-161">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|<span data-ttu-id="4dbdc-162">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-162">x</span></span>|<span data-ttu-id="4dbdc-163">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-163">x</span></span>|<span data-ttu-id="4dbdc-164">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-164">x</span></span>|
|[<span data-ttu-id="4dbdc-165">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="4dbdc-165">SupportUrl</span></span>](supporturl.md)|<span data-ttu-id="4dbdc-166">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-166">x</span></span>|<span data-ttu-id="4dbdc-167">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-167">x</span></span>|<span data-ttu-id="4dbdc-168">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-168">x</span></span>|
|[<span data-ttu-id="4dbdc-169">AppDomains</span><span class="sxs-lookup"><span data-stu-id="4dbdc-169">AppDomains</span></span>](appdomains.md)|<span data-ttu-id="4dbdc-170">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-170">x</span></span>|<span data-ttu-id="4dbdc-171">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-171">x</span></span>|<span data-ttu-id="4dbdc-172">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-172">x</span></span>|
|[<span data-ttu-id="4dbdc-173">Hosts</span><span class="sxs-lookup"><span data-stu-id="4dbdc-173">Hosts</span></span>](hosts.md)|<span data-ttu-id="4dbdc-174">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-174">x</span></span>|<span data-ttu-id="4dbdc-175">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-175">x</span></span>|<span data-ttu-id="4dbdc-176">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-176">x</span></span>|
|[<span data-ttu-id="4dbdc-177">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4dbdc-177">Requirements</span></span>](requirements.md)|<span data-ttu-id="4dbdc-178">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-178">x</span></span>|<span data-ttu-id="4dbdc-179">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-179">x</span></span>|<span data-ttu-id="4dbdc-180">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-180">x</span></span>|
|[<span data-ttu-id="4dbdc-181">AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="4dbdc-181">AllowSnapshot</span></span>](allowsnapshot.md)|<span data-ttu-id="4dbdc-182">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-182">x</span></span>|||
|[<span data-ttu-id="4dbdc-183">Permissões</span><span class="sxs-lookup"><span data-stu-id="4dbdc-183">Permissions</span></span>](permissions.md)||<span data-ttu-id="4dbdc-184">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-184">x</span></span>||
|[<span data-ttu-id="4dbdc-185">DisableEntityHighlighting</span><span class="sxs-lookup"><span data-stu-id="4dbdc-185">DisableEntityHighlighting</span></span>](disableentityhighlighting.md)||<span data-ttu-id="4dbdc-186">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-186">x</span></span>||
|[<span data-ttu-id="4dbdc-187">Dictionary</span><span class="sxs-lookup"><span data-stu-id="4dbdc-187">Dictionary</span></span>](dictionary.md)|||<span data-ttu-id="4dbdc-188">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-188">x</span></span>|
|[<span data-ttu-id="4dbdc-189">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="4dbdc-189">VersionOverrides</span></span>](versionoverrides.md)|<span data-ttu-id="4dbdc-190">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-190">x</span></span>|<span data-ttu-id="4dbdc-191">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-191">x</span></span>|<span data-ttu-id="4dbdc-192">x</span><span class="sxs-lookup"><span data-stu-id="4dbdc-192">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="4dbdc-193">Atributos</span><span class="sxs-lookup"><span data-stu-id="4dbdc-193">Attributes</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="4dbdc-194">xmlns</span><span class="sxs-lookup"><span data-stu-id="4dbdc-194">xmlns</span></span>|<span data-ttu-id="4dbdc-p101">Define o namespace do manifesto do Suplemento do Office e o esquema da versão. Esse atributo deve ser sempre definido como `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span><span class="sxs-lookup"><span data-stu-id="4dbdc-p101">Defines the Office Add-in manifest namespace and schema version. This attribute should always be set to  `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span></span>|
|<span data-ttu-id="4dbdc-197">xmlns: xsi</span><span class="sxs-lookup"><span data-stu-id="4dbdc-197">xmlns:xsi</span></span>|<span data-ttu-id="4dbdc-p102">Define a instância XMLSchema. Esse atributo deve ser sempre definido como `"http://www.w3.org/2001/XMLSchema-instance"`</span><span class="sxs-lookup"><span data-stu-id="4dbdc-p102">Defines the XMLSchema instance. This attribute should always be set to  `"http://www.w3.org/2001/XMLSchema-instance"`</span></span>|
|<span data-ttu-id="4dbdc-200">xsi:type</span><span class="sxs-lookup"><span data-stu-id="4dbdc-200">xsi:type</span></span>|<span data-ttu-id="4dbdc-p103">Define o tipo de Suplemento do Office. Esse atributo deve ser definido como um destes: `"ContentApp"`, `"MailApp"` ou `"TaskPaneApp"`</span><span class="sxs-lookup"><span data-stu-id="4dbdc-p103">Defines the kind of Office Add-in. This attribute should be set to one of:  `"ContentApp"`,  `"MailApp"`, or  `"TaskPaneApp"`</span></span>|
