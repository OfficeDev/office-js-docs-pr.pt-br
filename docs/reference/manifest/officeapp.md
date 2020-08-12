---
title: Elemento OfficeApp no arquivo de manifesto
description: O elemento OfficeApp é o elemento raiz de um manifesto de suplemento do Office.
ms.date: 02/04/2020
localization_priority: Normal
ms.openlocfilehash: 770c764db6d8d7d1d2e870e48437de7c8f887101
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641456"
---
# <a name="officeapp-element"></a><span data-ttu-id="93758-103">Elemento OfficeApp</span><span class="sxs-lookup"><span data-stu-id="93758-103">OfficeApp element</span></span>

<span data-ttu-id="93758-104">O elemento raiz no manifesto de um Suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="93758-104">The root element in the manifest of an Office Add-in.</span></span>

<span data-ttu-id="93758-105">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="93758-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="93758-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="93758-106">Syntax</span></span>

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="93758-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="93758-107">Contained in</span></span>

 <span data-ttu-id="93758-108">_none_</span><span class="sxs-lookup"><span data-stu-id="93758-108">_none_</span></span>

## <a name="must-contain"></a><span data-ttu-id="93758-109">Deve conter</span><span class="sxs-lookup"><span data-stu-id="93758-109">Must contain</span></span>

|<span data-ttu-id="93758-110">Elemento</span><span class="sxs-lookup"><span data-stu-id="93758-110">Element</span></span>|<span data-ttu-id="93758-111">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="93758-111">Content</span></span>|<span data-ttu-id="93758-112">Email</span><span class="sxs-lookup"><span data-stu-id="93758-112">Mail</span></span>|<span data-ttu-id="93758-113">TaskPane</span><span class="sxs-lookup"><span data-stu-id="93758-113">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="93758-114">Id</span><span class="sxs-lookup"><span data-stu-id="93758-114">Id</span></span>](id.md)|<span data-ttu-id="93758-115">x</span><span class="sxs-lookup"><span data-stu-id="93758-115">x</span></span>|<span data-ttu-id="93758-116">x</span><span class="sxs-lookup"><span data-stu-id="93758-116">x</span></span>|<span data-ttu-id="93758-117">x</span><span class="sxs-lookup"><span data-stu-id="93758-117">x</span></span>|
|[<span data-ttu-id="93758-118">Version</span><span class="sxs-lookup"><span data-stu-id="93758-118">Version</span></span>](version.md)|<span data-ttu-id="93758-119">x</span><span class="sxs-lookup"><span data-stu-id="93758-119">x</span></span>|<span data-ttu-id="93758-120">x</span><span class="sxs-lookup"><span data-stu-id="93758-120">x</span></span>|<span data-ttu-id="93758-121">x</span><span class="sxs-lookup"><span data-stu-id="93758-121">x</span></span>|
|[<span data-ttu-id="93758-122">ProviderName</span><span class="sxs-lookup"><span data-stu-id="93758-122">ProviderName</span></span>](providername.md)|<span data-ttu-id="93758-123">x</span><span class="sxs-lookup"><span data-stu-id="93758-123">x</span></span>|<span data-ttu-id="93758-124">x</span><span class="sxs-lookup"><span data-stu-id="93758-124">x</span></span>|<span data-ttu-id="93758-125">x</span><span class="sxs-lookup"><span data-stu-id="93758-125">x</span></span>|
|[<span data-ttu-id="93758-126">DefaultLocale</span><span class="sxs-lookup"><span data-stu-id="93758-126">DefaultLocale</span></span>](defaultlocale.md)|<span data-ttu-id="93758-127">x</span><span class="sxs-lookup"><span data-stu-id="93758-127">x</span></span>|<span data-ttu-id="93758-128">x</span><span class="sxs-lookup"><span data-stu-id="93758-128">x</span></span>|<span data-ttu-id="93758-129">x</span><span class="sxs-lookup"><span data-stu-id="93758-129">x</span></span>|
|[<span data-ttu-id="93758-130">DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="93758-130">DefaultSettings</span></span>](defaultsettings.md)|<span data-ttu-id="93758-131">x</span><span class="sxs-lookup"><span data-stu-id="93758-131">x</span></span>||<span data-ttu-id="93758-132">x</span><span class="sxs-lookup"><span data-stu-id="93758-132">x</span></span>|
|[<span data-ttu-id="93758-133">DisplayName</span><span class="sxs-lookup"><span data-stu-id="93758-133">DisplayName</span></span>](displayname.md)|<span data-ttu-id="93758-134">x</span><span class="sxs-lookup"><span data-stu-id="93758-134">x</span></span>|<span data-ttu-id="93758-135">x</span><span class="sxs-lookup"><span data-stu-id="93758-135">x</span></span>|<span data-ttu-id="93758-136">x</span><span class="sxs-lookup"><span data-stu-id="93758-136">x</span></span>|
|[<span data-ttu-id="93758-137">Descrição</span><span class="sxs-lookup"><span data-stu-id="93758-137">Description</span></span>](description.md)|<span data-ttu-id="93758-138">x</span><span class="sxs-lookup"><span data-stu-id="93758-138">x</span></span>|<span data-ttu-id="93758-139">x</span><span class="sxs-lookup"><span data-stu-id="93758-139">x</span></span>|<span data-ttu-id="93758-140">x</span><span class="sxs-lookup"><span data-stu-id="93758-140">x</span></span>|
|[<span data-ttu-id="93758-141">FormSettings</span><span class="sxs-lookup"><span data-stu-id="93758-141">FormSettings</span></span>](formsettings.md)||<span data-ttu-id="93758-142">x</span><span class="sxs-lookup"><span data-stu-id="93758-142">x</span></span>||
|[<span data-ttu-id="93758-143">Permissões</span><span class="sxs-lookup"><span data-stu-id="93758-143">Permissions</span></span>](permissions.md)|<span data-ttu-id="93758-144">x</span><span class="sxs-lookup"><span data-stu-id="93758-144">x</span></span>||<span data-ttu-id="93758-145">x</span><span class="sxs-lookup"><span data-stu-id="93758-145">x</span></span>|
|[<span data-ttu-id="93758-146">Rule</span><span class="sxs-lookup"><span data-stu-id="93758-146">Rule</span></span>](rule.md)||<span data-ttu-id="93758-147">x</span><span class="sxs-lookup"><span data-stu-id="93758-147">x</span></span>||

## <a name="can-contain"></a><span data-ttu-id="93758-148">Pode conter</span><span class="sxs-lookup"><span data-stu-id="93758-148">Can contain</span></span>

|<span data-ttu-id="93758-149">Elemento</span><span class="sxs-lookup"><span data-stu-id="93758-149">Element</span></span>|<span data-ttu-id="93758-150">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="93758-150">Content</span></span>|<span data-ttu-id="93758-151">Email</span><span class="sxs-lookup"><span data-stu-id="93758-151">Mail</span></span>|<span data-ttu-id="93758-152">TaskPane</span><span class="sxs-lookup"><span data-stu-id="93758-152">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="93758-153">AlternateId</span><span class="sxs-lookup"><span data-stu-id="93758-153">AlternateId</span></span>](alternateid.md)|<span data-ttu-id="93758-154">x</span><span class="sxs-lookup"><span data-stu-id="93758-154">x</span></span>|<span data-ttu-id="93758-155">x</span><span class="sxs-lookup"><span data-stu-id="93758-155">x</span></span>|<span data-ttu-id="93758-156">x</span><span class="sxs-lookup"><span data-stu-id="93758-156">x</span></span>|
|[<span data-ttu-id="93758-157">IconUrl</span><span class="sxs-lookup"><span data-stu-id="93758-157">IconUrl</span></span>](iconurl.md)|<span data-ttu-id="93758-158">x</span><span class="sxs-lookup"><span data-stu-id="93758-158">x</span></span>|<span data-ttu-id="93758-159">x</span><span class="sxs-lookup"><span data-stu-id="93758-159">x</span></span>|<span data-ttu-id="93758-160">x</span><span class="sxs-lookup"><span data-stu-id="93758-160">x</span></span>|
|[<span data-ttu-id="93758-161">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="93758-161">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|<span data-ttu-id="93758-162">x</span><span class="sxs-lookup"><span data-stu-id="93758-162">x</span></span>|<span data-ttu-id="93758-163">x</span><span class="sxs-lookup"><span data-stu-id="93758-163">x</span></span>|<span data-ttu-id="93758-164">x</span><span class="sxs-lookup"><span data-stu-id="93758-164">x</span></span>|
|[<span data-ttu-id="93758-165">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="93758-165">SupportUrl</span></span>](supporturl.md)|<span data-ttu-id="93758-166">x</span><span class="sxs-lookup"><span data-stu-id="93758-166">x</span></span>|<span data-ttu-id="93758-167">x</span><span class="sxs-lookup"><span data-stu-id="93758-167">x</span></span>|<span data-ttu-id="93758-168">x</span><span class="sxs-lookup"><span data-stu-id="93758-168">x</span></span>|
|[<span data-ttu-id="93758-169">AppDomains</span><span class="sxs-lookup"><span data-stu-id="93758-169">AppDomains</span></span>](appdomains.md)|<span data-ttu-id="93758-170">x</span><span class="sxs-lookup"><span data-stu-id="93758-170">x</span></span>|<span data-ttu-id="93758-171">x</span><span class="sxs-lookup"><span data-stu-id="93758-171">x</span></span>|<span data-ttu-id="93758-172">x</span><span class="sxs-lookup"><span data-stu-id="93758-172">x</span></span>|
|[<span data-ttu-id="93758-173">Hosts</span><span class="sxs-lookup"><span data-stu-id="93758-173">Hosts</span></span>](hosts.md)|<span data-ttu-id="93758-174">x</span><span class="sxs-lookup"><span data-stu-id="93758-174">x</span></span>|<span data-ttu-id="93758-175">x</span><span class="sxs-lookup"><span data-stu-id="93758-175">x</span></span>|<span data-ttu-id="93758-176">x</span><span class="sxs-lookup"><span data-stu-id="93758-176">x</span></span>|
|[<span data-ttu-id="93758-177">Requisitos</span><span class="sxs-lookup"><span data-stu-id="93758-177">Requirements</span></span>](requirements.md)|<span data-ttu-id="93758-178">x</span><span class="sxs-lookup"><span data-stu-id="93758-178">x</span></span>|<span data-ttu-id="93758-179">x</span><span class="sxs-lookup"><span data-stu-id="93758-179">x</span></span>|<span data-ttu-id="93758-180">x</span><span class="sxs-lookup"><span data-stu-id="93758-180">x</span></span>|
|[<span data-ttu-id="93758-181">AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="93758-181">AllowSnapshot</span></span>](allowsnapshot.md)|<span data-ttu-id="93758-182">x</span><span class="sxs-lookup"><span data-stu-id="93758-182">x</span></span>|||
|[<span data-ttu-id="93758-183">Permissões</span><span class="sxs-lookup"><span data-stu-id="93758-183">Permissions</span></span>](permissions.md)||<span data-ttu-id="93758-184">x</span><span class="sxs-lookup"><span data-stu-id="93758-184">x</span></span>||
|[<span data-ttu-id="93758-185">DisableEntityHighlighting</span><span class="sxs-lookup"><span data-stu-id="93758-185">DisableEntityHighlighting</span></span>](disableentityhighlighting.md)||<span data-ttu-id="93758-186">x</span><span class="sxs-lookup"><span data-stu-id="93758-186">x</span></span>||
|[<span data-ttu-id="93758-187">Dictionary</span><span class="sxs-lookup"><span data-stu-id="93758-187">Dictionary</span></span>](dictionary.md)|||<span data-ttu-id="93758-188">x</span><span class="sxs-lookup"><span data-stu-id="93758-188">x</span></span>|
|[<span data-ttu-id="93758-189">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="93758-189">VersionOverrides</span></span>](versionoverrides.md)|<span data-ttu-id="93758-190">x</span><span class="sxs-lookup"><span data-stu-id="93758-190">x</span></span>|<span data-ttu-id="93758-191">x</span><span class="sxs-lookup"><span data-stu-id="93758-191">x</span></span>|<span data-ttu-id="93758-192">x</span><span class="sxs-lookup"><span data-stu-id="93758-192">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="93758-193">Atributos</span><span class="sxs-lookup"><span data-stu-id="93758-193">Attributes</span></span>

|<span data-ttu-id="93758-194">Atributo</span><span class="sxs-lookup"><span data-stu-id="93758-194">Attribute</span></span>|<span data-ttu-id="93758-195">Descrição</span><span class="sxs-lookup"><span data-stu-id="93758-195">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="93758-196">xmlns</span><span class="sxs-lookup"><span data-stu-id="93758-196">xmlns</span></span>|<span data-ttu-id="93758-p101">Define o namespace do manifesto do Suplemento do Office e o esquema da versão. Esse atributo deve ser sempre definido como `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span><span class="sxs-lookup"><span data-stu-id="93758-p101">Defines the Office Add-in manifest namespace and schema version. This attribute should always be set to  `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span></span>|
|<span data-ttu-id="93758-199">xmlns: xsi</span><span class="sxs-lookup"><span data-stu-id="93758-199">xmlns:xsi</span></span>|<span data-ttu-id="93758-p102">Define a instância XMLSchema. Esse atributo deve ser sempre definido como `"http://www.w3.org/2001/XMLSchema-instance"`</span><span class="sxs-lookup"><span data-stu-id="93758-p102">Defines the XMLSchema instance. This attribute should always be set to  `"http://www.w3.org/2001/XMLSchema-instance"`</span></span>|
|<span data-ttu-id="93758-202">xsi:type</span><span class="sxs-lookup"><span data-stu-id="93758-202">xsi:type</span></span>|<span data-ttu-id="93758-p103">Define o tipo de Suplemento do Office. Esse atributo deve ser definido como um destes: `"ContentApp"`, `"MailApp"` ou `"TaskPaneApp"`</span><span class="sxs-lookup"><span data-stu-id="93758-p103">Defines the kind of Office Add-in. This attribute should be set to one of:  `"ContentApp"`,  `"MailApp"`, or  `"TaskPaneApp"`</span></span>|
