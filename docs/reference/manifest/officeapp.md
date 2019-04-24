---
title: Elemento OfficeApp no arquivo de manifesto
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 86f38ab77e98bb01370e40c8ada38bae171e0c2d
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450454"
---
# <a name="officeapp-element"></a><span data-ttu-id="b96b2-102">Elemento OfficeApp</span><span class="sxs-lookup"><span data-stu-id="b96b2-102">OfficeApp element</span></span>

<span data-ttu-id="b96b2-103">O elemento raiz no manifesto de um Suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="b96b2-103">The root element in the manifest of an Office Add-in.</span></span>

<span data-ttu-id="b96b2-104">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="b96b2-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="b96b2-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="b96b2-105">Syntax</span></span>

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="b96b2-106">Contido em</span><span class="sxs-lookup"><span data-stu-id="b96b2-106">Contained in</span></span>

 <span data-ttu-id="b96b2-107">_none_</span><span class="sxs-lookup"><span data-stu-id="b96b2-107">_none_</span></span>

## <a name="must-contain"></a><span data-ttu-id="b96b2-108">Deve conter</span><span class="sxs-lookup"><span data-stu-id="b96b2-108">Must contain</span></span>

|<span data-ttu-id="b96b2-109">**Element**</span><span class="sxs-lookup"><span data-stu-id="b96b2-109">**Element**</span></span>|<span data-ttu-id="b96b2-110">**Content**</span><span class="sxs-lookup"><span data-stu-id="b96b2-110">**Content**</span></span>|<span data-ttu-id="b96b2-111">**Email**</span><span class="sxs-lookup"><span data-stu-id="b96b2-111">**Mail**</span></span>|<span data-ttu-id="b96b2-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="b96b2-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="b96b2-113">Id</span><span class="sxs-lookup"><span data-stu-id="b96b2-113">Id</span></span>](id.md)|<span data-ttu-id="b96b2-114">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-114">x</span></span>|<span data-ttu-id="b96b2-115">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-115">x</span></span>|<span data-ttu-id="b96b2-116">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-116">x</span></span>|
|[<span data-ttu-id="b96b2-117">Version</span><span class="sxs-lookup"><span data-stu-id="b96b2-117">Version</span></span>](version.md)|<span data-ttu-id="b96b2-118">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-118">x</span></span>|<span data-ttu-id="b96b2-119">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-119">x</span></span>|<span data-ttu-id="b96b2-120">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-120">x</span></span>|
|[<span data-ttu-id="b96b2-121">ProviderName</span><span class="sxs-lookup"><span data-stu-id="b96b2-121">ProviderName</span></span>](providername.md)|<span data-ttu-id="b96b2-122">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-122">x</span></span>|<span data-ttu-id="b96b2-123">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-123">x</span></span>|<span data-ttu-id="b96b2-124">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-124">x</span></span>|
|[<span data-ttu-id="b96b2-125">DefaultLocale</span><span class="sxs-lookup"><span data-stu-id="b96b2-125">DefaultLocale</span></span>](defaultlocale.md)|<span data-ttu-id="b96b2-126">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-126">x</span></span>|<span data-ttu-id="b96b2-127">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-127">x</span></span>|<span data-ttu-id="b96b2-128">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-128">x</span></span>|
|[<span data-ttu-id="b96b2-129">DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="b96b2-129">DefaultSettings</span></span>](defaultsettings.md)|<span data-ttu-id="b96b2-130">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-130">x</span></span>||<span data-ttu-id="b96b2-131">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-131">x</span></span>|
|[<span data-ttu-id="b96b2-132">DisplayName</span><span class="sxs-lookup"><span data-stu-id="b96b2-132">DisplayName</span></span>](displayname.md)|<span data-ttu-id="b96b2-133">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-133">x</span></span>|<span data-ttu-id="b96b2-134">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-134">x</span></span>|<span data-ttu-id="b96b2-135">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-135">x</span></span>|
|[<span data-ttu-id="b96b2-136">Descrição</span><span class="sxs-lookup"><span data-stu-id="b96b2-136">Description</span></span>](description.md)|<span data-ttu-id="b96b2-137">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-137">x</span></span>|<span data-ttu-id="b96b2-138">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-138">x</span></span>|<span data-ttu-id="b96b2-139">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-139">x</span></span>|
|[<span data-ttu-id="b96b2-140">FormSettings</span><span class="sxs-lookup"><span data-stu-id="b96b2-140">FormSettings</span></span>](formsettings.md)||<span data-ttu-id="b96b2-141">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-141">x</span></span>||
|[<span data-ttu-id="b96b2-142">Permissões</span><span class="sxs-lookup"><span data-stu-id="b96b2-142">Permissions</span></span>](permissions.md)|<span data-ttu-id="b96b2-143">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-143">x</span></span>||<span data-ttu-id="b96b2-144">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-144">x</span></span>|
|[<span data-ttu-id="b96b2-145">Rule</span><span class="sxs-lookup"><span data-stu-id="b96b2-145">Rule</span></span>](rule.md)||<span data-ttu-id="b96b2-146">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-146">x</span></span>||

## <a name="can-contain"></a><span data-ttu-id="b96b2-147">Pode conter</span><span class="sxs-lookup"><span data-stu-id="b96b2-147">Can contain</span></span>

|<span data-ttu-id="b96b2-148">**Elemento**</span><span class="sxs-lookup"><span data-stu-id="b96b2-148">**Element**</span></span>|<span data-ttu-id="b96b2-149">**Content**</span><span class="sxs-lookup"><span data-stu-id="b96b2-149">**Content**</span></span>|<span data-ttu-id="b96b2-150">**Email**</span><span class="sxs-lookup"><span data-stu-id="b96b2-150">**Mail**</span></span>|<span data-ttu-id="b96b2-151">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="b96b2-151">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="b96b2-152">AlternateId</span><span class="sxs-lookup"><span data-stu-id="b96b2-152">AlternateId</span></span>](alternateid.md)|<span data-ttu-id="b96b2-153">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-153">x</span></span>|<span data-ttu-id="b96b2-154">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-154">x</span></span>|<span data-ttu-id="b96b2-155">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-155">x</span></span>|
|[<span data-ttu-id="b96b2-156">IconUrl</span><span class="sxs-lookup"><span data-stu-id="b96b2-156">IconUrl</span></span>](iconurl.md)|<span data-ttu-id="b96b2-157">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-157">x</span></span>|<span data-ttu-id="b96b2-158">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-158">x</span></span>|<span data-ttu-id="b96b2-159">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-159">x</span></span>|
|[<span data-ttu-id="b96b2-160">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="b96b2-160">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|<span data-ttu-id="b96b2-161">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-161">x</span></span>|<span data-ttu-id="b96b2-162">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-162">x</span></span>|<span data-ttu-id="b96b2-163">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-163">x</span></span>|
|[<span data-ttu-id="b96b2-164">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="b96b2-164">SupportUrl</span></span>](supporturl.md)|<span data-ttu-id="b96b2-165">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-165">x</span></span>|<span data-ttu-id="b96b2-166">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-166">x</span></span>|<span data-ttu-id="b96b2-167">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-167">x</span></span>|
|[<span data-ttu-id="b96b2-168">AppDomains</span><span class="sxs-lookup"><span data-stu-id="b96b2-168">AppDomains</span></span>](appdomains.md)|<span data-ttu-id="b96b2-169">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-169">x</span></span>|<span data-ttu-id="b96b2-170">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-170">x</span></span>|<span data-ttu-id="b96b2-171">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-171">x</span></span>|
|[<span data-ttu-id="b96b2-172">Hosts</span><span class="sxs-lookup"><span data-stu-id="b96b2-172">Hosts</span></span>](hosts.md)|<span data-ttu-id="b96b2-173">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-173">x</span></span>|<span data-ttu-id="b96b2-174">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-174">x</span></span>|<span data-ttu-id="b96b2-175">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-175">x</span></span>|
|[<span data-ttu-id="b96b2-176">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b96b2-176">Requirements</span></span>](requirements.md)|<span data-ttu-id="b96b2-177">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-177">x</span></span>|<span data-ttu-id="b96b2-178">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-178">x</span></span>|<span data-ttu-id="b96b2-179">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-179">x</span></span>|
|[<span data-ttu-id="b96b2-180">AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="b96b2-180">AllowSnapshot</span></span>](allowsnapshot.md)|<span data-ttu-id="b96b2-181">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-181">x</span></span>|||
|[<span data-ttu-id="b96b2-182">Permissões</span><span class="sxs-lookup"><span data-stu-id="b96b2-182">Permissions</span></span>](permissions.md)||<span data-ttu-id="b96b2-183">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-183">x</span></span>||
|[<span data-ttu-id="b96b2-184">DisableEntityHighlighting</span><span class="sxs-lookup"><span data-stu-id="b96b2-184">DisableEntityHighlighting</span></span>](disableentityhighlighting.md)||<span data-ttu-id="b96b2-185">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-185">x</span></span>||
|[<span data-ttu-id="b96b2-186">Dictionary</span><span class="sxs-lookup"><span data-stu-id="b96b2-186">Dictionary</span></span>](dictionary.md)|||<span data-ttu-id="b96b2-187">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-187">x</span></span>|
|[<span data-ttu-id="b96b2-188">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="b96b2-188">VersionOverrides</span></span>](versionoverrides.md)||<span data-ttu-id="b96b2-189">x</span><span class="sxs-lookup"><span data-stu-id="b96b2-189">x</span></span>||

## <a name="attributes"></a><span data-ttu-id="b96b2-190">Atributos</span><span class="sxs-lookup"><span data-stu-id="b96b2-190">Attributes</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="b96b2-191">xmlns</span><span class="sxs-lookup"><span data-stu-id="b96b2-191">xmlns</span></span>|<span data-ttu-id="b96b2-p101">Define o namespace do manifesto do Suplemento do Office e o esquema da versão. Esse atributo deve ser sempre definido como `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span><span class="sxs-lookup"><span data-stu-id="b96b2-p101">Defines the Office Add-in manifest namespace and schema version. This attribute should always be set to  `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span></span>|
|<span data-ttu-id="b96b2-194">xmlns: xsi</span><span class="sxs-lookup"><span data-stu-id="b96b2-194">xmlns:xsi</span></span>|<span data-ttu-id="b96b2-p102">Define a instância XMLSchema. Esse atributo deve ser sempre definido como `"http://www.w3.org/2001/XMLSchema-instance"`</span><span class="sxs-lookup"><span data-stu-id="b96b2-p102">Defines the XMLSchema instance. This attribute should always be set to  `"http://www.w3.org/2001/XMLSchema-instance"`</span></span>|
|<span data-ttu-id="b96b2-197">xsi:type</span><span class="sxs-lookup"><span data-stu-id="b96b2-197">xsi:type</span></span>|<span data-ttu-id="b96b2-p103">Define o tipo de Suplemento do Office. Esse atributo deve ser definido como um destes: `"ContentApp"`, `"MailApp"` ou `"TaskPaneApp"`</span><span class="sxs-lookup"><span data-stu-id="b96b2-p103">Defines the kind of Office Add-in. This attribute should be set to one of:  `"ContentApp"`,  `"MailApp"`, or  `"TaskPaneApp"`</span></span>|
