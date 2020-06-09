---
title: Elemento OfficeApp no arquivo de manifesto
description: O elemento OfficeApp é o elemento raiz de um manifesto de suplemento do Office.
ms.date: 02/04/2020
localization_priority: Normal
ms.openlocfilehash: b6f3102a97794a19366b06734789e01fc4bc4f9d
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611523"
---
# <a name="officeapp-element"></a><span data-ttu-id="02bfe-103">Elemento OfficeApp</span><span class="sxs-lookup"><span data-stu-id="02bfe-103">OfficeApp element</span></span>

<span data-ttu-id="02bfe-104">O elemento raiz no manifesto de um Suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="02bfe-104">The root element in the manifest of an Office Add-in.</span></span>

<span data-ttu-id="02bfe-105">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="02bfe-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="02bfe-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="02bfe-106">Syntax</span></span>

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="02bfe-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="02bfe-107">Contained in</span></span>

 <span data-ttu-id="02bfe-108">_none_</span><span class="sxs-lookup"><span data-stu-id="02bfe-108">_none_</span></span>

## <a name="must-contain"></a><span data-ttu-id="02bfe-109">Deve conter</span><span class="sxs-lookup"><span data-stu-id="02bfe-109">Must contain</span></span>

|<span data-ttu-id="02bfe-110">**Element**</span><span class="sxs-lookup"><span data-stu-id="02bfe-110">**Element**</span></span>|<span data-ttu-id="02bfe-111">**Content**</span><span class="sxs-lookup"><span data-stu-id="02bfe-111">**Content**</span></span>|<span data-ttu-id="02bfe-112">**Email**</span><span class="sxs-lookup"><span data-stu-id="02bfe-112">**Mail**</span></span>|<span data-ttu-id="02bfe-113">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="02bfe-113">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="02bfe-114">Id</span><span class="sxs-lookup"><span data-stu-id="02bfe-114">Id</span></span>](id.md)|<span data-ttu-id="02bfe-115">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-115">x</span></span>|<span data-ttu-id="02bfe-116">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-116">x</span></span>|<span data-ttu-id="02bfe-117">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-117">x</span></span>|
|[<span data-ttu-id="02bfe-118">Version</span><span class="sxs-lookup"><span data-stu-id="02bfe-118">Version</span></span>](version.md)|<span data-ttu-id="02bfe-119">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-119">x</span></span>|<span data-ttu-id="02bfe-120">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-120">x</span></span>|<span data-ttu-id="02bfe-121">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-121">x</span></span>|
|[<span data-ttu-id="02bfe-122">ProviderName</span><span class="sxs-lookup"><span data-stu-id="02bfe-122">ProviderName</span></span>](providername.md)|<span data-ttu-id="02bfe-123">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-123">x</span></span>|<span data-ttu-id="02bfe-124">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-124">x</span></span>|<span data-ttu-id="02bfe-125">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-125">x</span></span>|
|[<span data-ttu-id="02bfe-126">DefaultLocale</span><span class="sxs-lookup"><span data-stu-id="02bfe-126">DefaultLocale</span></span>](defaultlocale.md)|<span data-ttu-id="02bfe-127">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-127">x</span></span>|<span data-ttu-id="02bfe-128">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-128">x</span></span>|<span data-ttu-id="02bfe-129">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-129">x</span></span>|
|[<span data-ttu-id="02bfe-130">DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="02bfe-130">DefaultSettings</span></span>](defaultsettings.md)|<span data-ttu-id="02bfe-131">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-131">x</span></span>||<span data-ttu-id="02bfe-132">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-132">x</span></span>|
|[<span data-ttu-id="02bfe-133">DisplayName</span><span class="sxs-lookup"><span data-stu-id="02bfe-133">DisplayName</span></span>](displayname.md)|<span data-ttu-id="02bfe-134">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-134">x</span></span>|<span data-ttu-id="02bfe-135">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-135">x</span></span>|<span data-ttu-id="02bfe-136">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-136">x</span></span>|
|[<span data-ttu-id="02bfe-137">Descrição</span><span class="sxs-lookup"><span data-stu-id="02bfe-137">Description</span></span>](description.md)|<span data-ttu-id="02bfe-138">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-138">x</span></span>|<span data-ttu-id="02bfe-139">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-139">x</span></span>|<span data-ttu-id="02bfe-140">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-140">x</span></span>|
|[<span data-ttu-id="02bfe-141">FormSettings</span><span class="sxs-lookup"><span data-stu-id="02bfe-141">FormSettings</span></span>](formsettings.md)||<span data-ttu-id="02bfe-142">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-142">x</span></span>||
|[<span data-ttu-id="02bfe-143">Permissões</span><span class="sxs-lookup"><span data-stu-id="02bfe-143">Permissions</span></span>](permissions.md)|<span data-ttu-id="02bfe-144">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-144">x</span></span>||<span data-ttu-id="02bfe-145">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-145">x</span></span>|
|[<span data-ttu-id="02bfe-146">Rule</span><span class="sxs-lookup"><span data-stu-id="02bfe-146">Rule</span></span>](rule.md)||<span data-ttu-id="02bfe-147">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-147">x</span></span>||

## <a name="can-contain"></a><span data-ttu-id="02bfe-148">Pode conter</span><span class="sxs-lookup"><span data-stu-id="02bfe-148">Can contain</span></span>

|<span data-ttu-id="02bfe-149">**Elemento**</span><span class="sxs-lookup"><span data-stu-id="02bfe-149">**Element**</span></span>|<span data-ttu-id="02bfe-150">**Content**</span><span class="sxs-lookup"><span data-stu-id="02bfe-150">**Content**</span></span>|<span data-ttu-id="02bfe-151">**Email**</span><span class="sxs-lookup"><span data-stu-id="02bfe-151">**Mail**</span></span>|<span data-ttu-id="02bfe-152">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="02bfe-152">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="02bfe-153">AlternateId</span><span class="sxs-lookup"><span data-stu-id="02bfe-153">AlternateId</span></span>](alternateid.md)|<span data-ttu-id="02bfe-154">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-154">x</span></span>|<span data-ttu-id="02bfe-155">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-155">x</span></span>|<span data-ttu-id="02bfe-156">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-156">x</span></span>|
|[<span data-ttu-id="02bfe-157">IconUrl</span><span class="sxs-lookup"><span data-stu-id="02bfe-157">IconUrl</span></span>](iconurl.md)|<span data-ttu-id="02bfe-158">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-158">x</span></span>|<span data-ttu-id="02bfe-159">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-159">x</span></span>|<span data-ttu-id="02bfe-160">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-160">x</span></span>|
|[<span data-ttu-id="02bfe-161">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="02bfe-161">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|<span data-ttu-id="02bfe-162">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-162">x</span></span>|<span data-ttu-id="02bfe-163">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-163">x</span></span>|<span data-ttu-id="02bfe-164">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-164">x</span></span>|
|[<span data-ttu-id="02bfe-165">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="02bfe-165">SupportUrl</span></span>](supporturl.md)|<span data-ttu-id="02bfe-166">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-166">x</span></span>|<span data-ttu-id="02bfe-167">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-167">x</span></span>|<span data-ttu-id="02bfe-168">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-168">x</span></span>|
|[<span data-ttu-id="02bfe-169">AppDomains</span><span class="sxs-lookup"><span data-stu-id="02bfe-169">AppDomains</span></span>](appdomains.md)|<span data-ttu-id="02bfe-170">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-170">x</span></span>|<span data-ttu-id="02bfe-171">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-171">x</span></span>|<span data-ttu-id="02bfe-172">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-172">x</span></span>|
|[<span data-ttu-id="02bfe-173">Hosts</span><span class="sxs-lookup"><span data-stu-id="02bfe-173">Hosts</span></span>](hosts.md)|<span data-ttu-id="02bfe-174">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-174">x</span></span>|<span data-ttu-id="02bfe-175">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-175">x</span></span>|<span data-ttu-id="02bfe-176">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-176">x</span></span>|
|[<span data-ttu-id="02bfe-177">Requisitos</span><span class="sxs-lookup"><span data-stu-id="02bfe-177">Requirements</span></span>](requirements.md)|<span data-ttu-id="02bfe-178">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-178">x</span></span>|<span data-ttu-id="02bfe-179">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-179">x</span></span>|<span data-ttu-id="02bfe-180">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-180">x</span></span>|
|[<span data-ttu-id="02bfe-181">AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="02bfe-181">AllowSnapshot</span></span>](allowsnapshot.md)|<span data-ttu-id="02bfe-182">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-182">x</span></span>|||
|[<span data-ttu-id="02bfe-183">Permissões</span><span class="sxs-lookup"><span data-stu-id="02bfe-183">Permissions</span></span>](permissions.md)||<span data-ttu-id="02bfe-184">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-184">x</span></span>||
|[<span data-ttu-id="02bfe-185">DisableEntityHighlighting</span><span class="sxs-lookup"><span data-stu-id="02bfe-185">DisableEntityHighlighting</span></span>](disableentityhighlighting.md)||<span data-ttu-id="02bfe-186">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-186">x</span></span>||
|[<span data-ttu-id="02bfe-187">Dictionary</span><span class="sxs-lookup"><span data-stu-id="02bfe-187">Dictionary</span></span>](dictionary.md)|||<span data-ttu-id="02bfe-188">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-188">x</span></span>|
|[<span data-ttu-id="02bfe-189">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="02bfe-189">VersionOverrides</span></span>](versionoverrides.md)|<span data-ttu-id="02bfe-190">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-190">x</span></span>|<span data-ttu-id="02bfe-191">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-191">x</span></span>|<span data-ttu-id="02bfe-192">x</span><span class="sxs-lookup"><span data-stu-id="02bfe-192">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="02bfe-193">Atributos</span><span class="sxs-lookup"><span data-stu-id="02bfe-193">Attributes</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="02bfe-194">xmlns</span><span class="sxs-lookup"><span data-stu-id="02bfe-194">xmlns</span></span>|<span data-ttu-id="02bfe-p101">Define o namespace do manifesto do Suplemento do Office e o esquema da versão. Esse atributo deve ser sempre definido como `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span><span class="sxs-lookup"><span data-stu-id="02bfe-p101">Defines the Office Add-in manifest namespace and schema version. This attribute should always be set to  `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span></span>|
|<span data-ttu-id="02bfe-197">xmlns: xsi</span><span class="sxs-lookup"><span data-stu-id="02bfe-197">xmlns:xsi</span></span>|<span data-ttu-id="02bfe-p102">Define a instância XMLSchema. Esse atributo deve ser sempre definido como `"http://www.w3.org/2001/XMLSchema-instance"`</span><span class="sxs-lookup"><span data-stu-id="02bfe-p102">Defines the XMLSchema instance. This attribute should always be set to  `"http://www.w3.org/2001/XMLSchema-instance"`</span></span>|
|<span data-ttu-id="02bfe-200">xsi:type</span><span class="sxs-lookup"><span data-stu-id="02bfe-200">xsi:type</span></span>|<span data-ttu-id="02bfe-p103">Define o tipo de Suplemento do Office. Esse atributo deve ser definido como um destes: `"ContentApp"`, `"MailApp"` ou `"TaskPaneApp"`</span><span class="sxs-lookup"><span data-stu-id="02bfe-p103">Defines the kind of Office Add-in. This attribute should be set to one of:  `"ContentApp"`,  `"MailApp"`, or  `"TaskPaneApp"`</span></span>|
