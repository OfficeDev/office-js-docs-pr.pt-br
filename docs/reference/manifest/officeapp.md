---
title: Elemento OfficeApp no arquivo de manifesto
description: O elemento OfficeApp é o elemento raiz de um manifesto de suplemento do Office.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: c5786343173d0e130df4b786f28a8689d573b6ca
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996316"
---
# <a name="officeapp-element"></a><span data-ttu-id="db55b-103">Elemento OfficeApp</span><span class="sxs-lookup"><span data-stu-id="db55b-103">OfficeApp element</span></span>

<span data-ttu-id="db55b-104">O elemento raiz no manifesto de um Suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="db55b-104">The root element in the manifest of an Office Add-in.</span></span>

<span data-ttu-id="db55b-105">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="db55b-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="db55b-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="db55b-106">Syntax</span></span>

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="db55b-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="db55b-107">Contained in</span></span>

 <span data-ttu-id="db55b-108">_none_</span><span class="sxs-lookup"><span data-stu-id="db55b-108">_none_</span></span>

## <a name="must-contain"></a><span data-ttu-id="db55b-109">Deve conter</span><span class="sxs-lookup"><span data-stu-id="db55b-109">Must contain</span></span>

|<span data-ttu-id="db55b-110">Elemento</span><span class="sxs-lookup"><span data-stu-id="db55b-110">Element</span></span>|<span data-ttu-id="db55b-111">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="db55b-111">Content</span></span>|<span data-ttu-id="db55b-112">Email</span><span class="sxs-lookup"><span data-stu-id="db55b-112">Mail</span></span>|<span data-ttu-id="db55b-113">TaskPane</span><span class="sxs-lookup"><span data-stu-id="db55b-113">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="db55b-114">Id</span><span class="sxs-lookup"><span data-stu-id="db55b-114">Id</span></span>](id.md)|<span data-ttu-id="db55b-115">x</span><span class="sxs-lookup"><span data-stu-id="db55b-115">x</span></span>|<span data-ttu-id="db55b-116">x</span><span class="sxs-lookup"><span data-stu-id="db55b-116">x</span></span>|<span data-ttu-id="db55b-117">x</span><span class="sxs-lookup"><span data-stu-id="db55b-117">x</span></span>|
|[<span data-ttu-id="db55b-118">Version</span><span class="sxs-lookup"><span data-stu-id="db55b-118">Version</span></span>](version.md)|<span data-ttu-id="db55b-119">x</span><span class="sxs-lookup"><span data-stu-id="db55b-119">x</span></span>|<span data-ttu-id="db55b-120">x</span><span class="sxs-lookup"><span data-stu-id="db55b-120">x</span></span>|<span data-ttu-id="db55b-121">x</span><span class="sxs-lookup"><span data-stu-id="db55b-121">x</span></span>|
|[<span data-ttu-id="db55b-122">ProviderName</span><span class="sxs-lookup"><span data-stu-id="db55b-122">ProviderName</span></span>](providername.md)|<span data-ttu-id="db55b-123">x</span><span class="sxs-lookup"><span data-stu-id="db55b-123">x</span></span>|<span data-ttu-id="db55b-124">x</span><span class="sxs-lookup"><span data-stu-id="db55b-124">x</span></span>|<span data-ttu-id="db55b-125">x</span><span class="sxs-lookup"><span data-stu-id="db55b-125">x</span></span>|
|[<span data-ttu-id="db55b-126">DefaultLocale</span><span class="sxs-lookup"><span data-stu-id="db55b-126">DefaultLocale</span></span>](defaultlocale.md)|<span data-ttu-id="db55b-127">x</span><span class="sxs-lookup"><span data-stu-id="db55b-127">x</span></span>|<span data-ttu-id="db55b-128">x</span><span class="sxs-lookup"><span data-stu-id="db55b-128">x</span></span>|<span data-ttu-id="db55b-129">x</span><span class="sxs-lookup"><span data-stu-id="db55b-129">x</span></span>|
|[<span data-ttu-id="db55b-130">DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="db55b-130">DefaultSettings</span></span>](defaultsettings.md)|<span data-ttu-id="db55b-131">x</span><span class="sxs-lookup"><span data-stu-id="db55b-131">x</span></span>||<span data-ttu-id="db55b-132">x</span><span class="sxs-lookup"><span data-stu-id="db55b-132">x</span></span>|
|[<span data-ttu-id="db55b-133">DisplayName</span><span class="sxs-lookup"><span data-stu-id="db55b-133">DisplayName</span></span>](displayname.md)|<span data-ttu-id="db55b-134">x</span><span class="sxs-lookup"><span data-stu-id="db55b-134">x</span></span>|<span data-ttu-id="db55b-135">x</span><span class="sxs-lookup"><span data-stu-id="db55b-135">x</span></span>|<span data-ttu-id="db55b-136">x</span><span class="sxs-lookup"><span data-stu-id="db55b-136">x</span></span>|
|[<span data-ttu-id="db55b-137">Descrição</span><span class="sxs-lookup"><span data-stu-id="db55b-137">Description</span></span>](description.md)|<span data-ttu-id="db55b-138">x</span><span class="sxs-lookup"><span data-stu-id="db55b-138">x</span></span>|<span data-ttu-id="db55b-139">x</span><span class="sxs-lookup"><span data-stu-id="db55b-139">x</span></span>|<span data-ttu-id="db55b-140">x</span><span class="sxs-lookup"><span data-stu-id="db55b-140">x</span></span>|
|[<span data-ttu-id="db55b-141">FormSettings</span><span class="sxs-lookup"><span data-stu-id="db55b-141">FormSettings</span></span>](formsettings.md)||<span data-ttu-id="db55b-142">x</span><span class="sxs-lookup"><span data-stu-id="db55b-142">x</span></span>||
|[<span data-ttu-id="db55b-143">Permissões</span><span class="sxs-lookup"><span data-stu-id="db55b-143">Permissions</span></span>](permissions.md)|<span data-ttu-id="db55b-144">x</span><span class="sxs-lookup"><span data-stu-id="db55b-144">x</span></span>||<span data-ttu-id="db55b-145">x</span><span class="sxs-lookup"><span data-stu-id="db55b-145">x</span></span>|
|[<span data-ttu-id="db55b-146">Rule</span><span class="sxs-lookup"><span data-stu-id="db55b-146">Rule</span></span>](rule.md)||<span data-ttu-id="db55b-147">x</span><span class="sxs-lookup"><span data-stu-id="db55b-147">x</span></span>||

## <a name="can-contain"></a><span data-ttu-id="db55b-148">Pode conter</span><span class="sxs-lookup"><span data-stu-id="db55b-148">Can contain</span></span>

|<span data-ttu-id="db55b-149">Elemento</span><span class="sxs-lookup"><span data-stu-id="db55b-149">Element</span></span>|<span data-ttu-id="db55b-150">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="db55b-150">Content</span></span>|<span data-ttu-id="db55b-151">Email</span><span class="sxs-lookup"><span data-stu-id="db55b-151">Mail</span></span>|<span data-ttu-id="db55b-152">TaskPane</span><span class="sxs-lookup"><span data-stu-id="db55b-152">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="db55b-153">AlternateId</span><span class="sxs-lookup"><span data-stu-id="db55b-153">AlternateId</span></span>](alternateid.md)|<span data-ttu-id="db55b-154">x</span><span class="sxs-lookup"><span data-stu-id="db55b-154">x</span></span>|<span data-ttu-id="db55b-155">x</span><span class="sxs-lookup"><span data-stu-id="db55b-155">x</span></span>|<span data-ttu-id="db55b-156">x</span><span class="sxs-lookup"><span data-stu-id="db55b-156">x</span></span>|
|[<span data-ttu-id="db55b-157">IconUrl</span><span class="sxs-lookup"><span data-stu-id="db55b-157">IconUrl</span></span>](iconurl.md)|<span data-ttu-id="db55b-158">x</span><span class="sxs-lookup"><span data-stu-id="db55b-158">x</span></span>|<span data-ttu-id="db55b-159">x</span><span class="sxs-lookup"><span data-stu-id="db55b-159">x</span></span>|<span data-ttu-id="db55b-160">x</span><span class="sxs-lookup"><span data-stu-id="db55b-160">x</span></span>|
|[<span data-ttu-id="db55b-161">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="db55b-161">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|<span data-ttu-id="db55b-162">x</span><span class="sxs-lookup"><span data-stu-id="db55b-162">x</span></span>|<span data-ttu-id="db55b-163">x</span><span class="sxs-lookup"><span data-stu-id="db55b-163">x</span></span>|<span data-ttu-id="db55b-164">x</span><span class="sxs-lookup"><span data-stu-id="db55b-164">x</span></span>|
|[<span data-ttu-id="db55b-165">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="db55b-165">SupportUrl</span></span>](supporturl.md)|<span data-ttu-id="db55b-166">x</span><span class="sxs-lookup"><span data-stu-id="db55b-166">x</span></span>|<span data-ttu-id="db55b-167">x</span><span class="sxs-lookup"><span data-stu-id="db55b-167">x</span></span>|<span data-ttu-id="db55b-168">x</span><span class="sxs-lookup"><span data-stu-id="db55b-168">x</span></span>|
|[<span data-ttu-id="db55b-169">AppDomains</span><span class="sxs-lookup"><span data-stu-id="db55b-169">AppDomains</span></span>](appdomains.md)|<span data-ttu-id="db55b-170">x</span><span class="sxs-lookup"><span data-stu-id="db55b-170">x</span></span>|<span data-ttu-id="db55b-171">x</span><span class="sxs-lookup"><span data-stu-id="db55b-171">x</span></span>|<span data-ttu-id="db55b-172">x</span><span class="sxs-lookup"><span data-stu-id="db55b-172">x</span></span>|
|[<span data-ttu-id="db55b-173">Hosts</span><span class="sxs-lookup"><span data-stu-id="db55b-173">Hosts</span></span>](hosts.md)|<span data-ttu-id="db55b-174">x</span><span class="sxs-lookup"><span data-stu-id="db55b-174">x</span></span>|<span data-ttu-id="db55b-175">x</span><span class="sxs-lookup"><span data-stu-id="db55b-175">x</span></span>|<span data-ttu-id="db55b-176">x</span><span class="sxs-lookup"><span data-stu-id="db55b-176">x</span></span>|
|[<span data-ttu-id="db55b-177">Requisitos</span><span class="sxs-lookup"><span data-stu-id="db55b-177">Requirements</span></span>](requirements.md)|<span data-ttu-id="db55b-178">x</span><span class="sxs-lookup"><span data-stu-id="db55b-178">x</span></span>|<span data-ttu-id="db55b-179">x</span><span class="sxs-lookup"><span data-stu-id="db55b-179">x</span></span>|<span data-ttu-id="db55b-180">x</span><span class="sxs-lookup"><span data-stu-id="db55b-180">x</span></span>|
|[<span data-ttu-id="db55b-181">AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="db55b-181">AllowSnapshot</span></span>](allowsnapshot.md)|<span data-ttu-id="db55b-182">x</span><span class="sxs-lookup"><span data-stu-id="db55b-182">x</span></span>|||
|[<span data-ttu-id="db55b-183">Permissões</span><span class="sxs-lookup"><span data-stu-id="db55b-183">Permissions</span></span>](permissions.md)||<span data-ttu-id="db55b-184">x</span><span class="sxs-lookup"><span data-stu-id="db55b-184">x</span></span>||
|[<span data-ttu-id="db55b-185">DisableEntityHighlighting</span><span class="sxs-lookup"><span data-stu-id="db55b-185">DisableEntityHighlighting</span></span>](disableentityhighlighting.md)||<span data-ttu-id="db55b-186">x</span><span class="sxs-lookup"><span data-stu-id="db55b-186">x</span></span>||
|[<span data-ttu-id="db55b-187">Dictionary</span><span class="sxs-lookup"><span data-stu-id="db55b-187">Dictionary</span></span>](dictionary.md)|||<span data-ttu-id="db55b-188">x</span><span class="sxs-lookup"><span data-stu-id="db55b-188">x</span></span>|
|[<span data-ttu-id="db55b-189">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="db55b-189">VersionOverrides</span></span>](versionoverrides.md)|<span data-ttu-id="db55b-190">x</span><span class="sxs-lookup"><span data-stu-id="db55b-190">x</span></span>|<span data-ttu-id="db55b-191">x</span><span class="sxs-lookup"><span data-stu-id="db55b-191">x</span></span>|<span data-ttu-id="db55b-192">x</span><span class="sxs-lookup"><span data-stu-id="db55b-192">x</span></span>|
|[<span data-ttu-id="db55b-193">ExtendedOverrides</span><span class="sxs-lookup"><span data-stu-id="db55b-193">ExtendedOverrides</span></span>](extendedoverrides.md)|||<span data-ttu-id="db55b-194">x</span><span class="sxs-lookup"><span data-stu-id="db55b-194">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="db55b-195">Atributos</span><span class="sxs-lookup"><span data-stu-id="db55b-195">Attributes</span></span>

|<span data-ttu-id="db55b-196">Atributo</span><span class="sxs-lookup"><span data-stu-id="db55b-196">Attribute</span></span>|<span data-ttu-id="db55b-197">Descrição</span><span class="sxs-lookup"><span data-stu-id="db55b-197">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="db55b-198">xmlns</span><span class="sxs-lookup"><span data-stu-id="db55b-198">xmlns</span></span>|<span data-ttu-id="db55b-p101">Define o namespace do manifesto do Suplemento do Office e o esquema da versão. Esse atributo deve ser sempre definido como `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span><span class="sxs-lookup"><span data-stu-id="db55b-p101">Defines the Office Add-in manifest namespace and schema version. This attribute should always be set to  `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span></span>|
|<span data-ttu-id="db55b-201">xmlns: xsi</span><span class="sxs-lookup"><span data-stu-id="db55b-201">xmlns:xsi</span></span>|<span data-ttu-id="db55b-p102">Define a instância XMLSchema. Esse atributo deve ser sempre definido como `"http://www.w3.org/2001/XMLSchema-instance"`</span><span class="sxs-lookup"><span data-stu-id="db55b-p102">Defines the XMLSchema instance. This attribute should always be set to  `"http://www.w3.org/2001/XMLSchema-instance"`</span></span>|
|<span data-ttu-id="db55b-204">xsi:type</span><span class="sxs-lookup"><span data-stu-id="db55b-204">xsi:type</span></span>|<span data-ttu-id="db55b-p103">Define o tipo de Suplemento do Office. Esse atributo deve ser definido como um destes: `"ContentApp"`, `"MailApp"` ou `"TaskPaneApp"`</span><span class="sxs-lookup"><span data-stu-id="db55b-p103">Defines the kind of Office Add-in. This attribute should be set to one of:  `"ContentApp"`,  `"MailApp"`, or  `"TaskPaneApp"`</span></span>|
