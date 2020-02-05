---
title: Elemento OfficeApp no arquivo de manifesto
description: ''
ms.date: 02/04/2020
localization_priority: Normal
ms.openlocfilehash: 080025e62a56421dff942792f99ee672ce1db69a
ms.sourcegitcommit: c1dbea577ae6183523fb663d364422d2adbc8bcf
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/05/2020
ms.locfileid: "41773576"
---
# <a name="officeapp-element"></a><span data-ttu-id="9c54b-102">Elemento OfficeApp</span><span class="sxs-lookup"><span data-stu-id="9c54b-102">OfficeApp element</span></span>

<span data-ttu-id="9c54b-103">O elemento raiz no manifesto de um Suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="9c54b-103">The root element in the manifest of an Office Add-in.</span></span>

<span data-ttu-id="9c54b-104">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="9c54b-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="9c54b-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="9c54b-105">Syntax</span></span>

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="9c54b-106">Contido em</span><span class="sxs-lookup"><span data-stu-id="9c54b-106">Contained in</span></span>

 <span data-ttu-id="9c54b-107">_none_</span><span class="sxs-lookup"><span data-stu-id="9c54b-107">_none_</span></span>

## <a name="must-contain"></a><span data-ttu-id="9c54b-108">Deve conter</span><span class="sxs-lookup"><span data-stu-id="9c54b-108">Must contain</span></span>

|<span data-ttu-id="9c54b-109">**Element**</span><span class="sxs-lookup"><span data-stu-id="9c54b-109">**Element**</span></span>|<span data-ttu-id="9c54b-110">**Content**</span><span class="sxs-lookup"><span data-stu-id="9c54b-110">**Content**</span></span>|<span data-ttu-id="9c54b-111">**Email**</span><span class="sxs-lookup"><span data-stu-id="9c54b-111">**Mail**</span></span>|<span data-ttu-id="9c54b-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="9c54b-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="9c54b-113">Id</span><span class="sxs-lookup"><span data-stu-id="9c54b-113">Id</span></span>](id.md)|<span data-ttu-id="9c54b-114">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-114">x</span></span>|<span data-ttu-id="9c54b-115">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-115">x</span></span>|<span data-ttu-id="9c54b-116">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-116">x</span></span>|
|[<span data-ttu-id="9c54b-117">Version</span><span class="sxs-lookup"><span data-stu-id="9c54b-117">Version</span></span>](version.md)|<span data-ttu-id="9c54b-118">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-118">x</span></span>|<span data-ttu-id="9c54b-119">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-119">x</span></span>|<span data-ttu-id="9c54b-120">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-120">x</span></span>|
|[<span data-ttu-id="9c54b-121">ProviderName</span><span class="sxs-lookup"><span data-stu-id="9c54b-121">ProviderName</span></span>](providername.md)|<span data-ttu-id="9c54b-122">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-122">x</span></span>|<span data-ttu-id="9c54b-123">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-123">x</span></span>|<span data-ttu-id="9c54b-124">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-124">x</span></span>|
|[<span data-ttu-id="9c54b-125">DefaultLocale</span><span class="sxs-lookup"><span data-stu-id="9c54b-125">DefaultLocale</span></span>](defaultlocale.md)|<span data-ttu-id="9c54b-126">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-126">x</span></span>|<span data-ttu-id="9c54b-127">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-127">x</span></span>|<span data-ttu-id="9c54b-128">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-128">x</span></span>|
|[<span data-ttu-id="9c54b-129">DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="9c54b-129">DefaultSettings</span></span>](defaultsettings.md)|<span data-ttu-id="9c54b-130">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-130">x</span></span>||<span data-ttu-id="9c54b-131">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-131">x</span></span>|
|[<span data-ttu-id="9c54b-132">DisplayName</span><span class="sxs-lookup"><span data-stu-id="9c54b-132">DisplayName</span></span>](displayname.md)|<span data-ttu-id="9c54b-133">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-133">x</span></span>|<span data-ttu-id="9c54b-134">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-134">x</span></span>|<span data-ttu-id="9c54b-135">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-135">x</span></span>|
|[<span data-ttu-id="9c54b-136">Descrição</span><span class="sxs-lookup"><span data-stu-id="9c54b-136">Description</span></span>](description.md)|<span data-ttu-id="9c54b-137">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-137">x</span></span>|<span data-ttu-id="9c54b-138">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-138">x</span></span>|<span data-ttu-id="9c54b-139">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-139">x</span></span>|
|[<span data-ttu-id="9c54b-140">FormSettings</span><span class="sxs-lookup"><span data-stu-id="9c54b-140">FormSettings</span></span>](formsettings.md)||<span data-ttu-id="9c54b-141">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-141">x</span></span>||
|[<span data-ttu-id="9c54b-142">Permissões</span><span class="sxs-lookup"><span data-stu-id="9c54b-142">Permissions</span></span>](permissions.md)|<span data-ttu-id="9c54b-143">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-143">x</span></span>||<span data-ttu-id="9c54b-144">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-144">x</span></span>|
|[<span data-ttu-id="9c54b-145">Rule</span><span class="sxs-lookup"><span data-stu-id="9c54b-145">Rule</span></span>](rule.md)||<span data-ttu-id="9c54b-146">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-146">x</span></span>||

## <a name="can-contain"></a><span data-ttu-id="9c54b-147">Pode conter</span><span class="sxs-lookup"><span data-stu-id="9c54b-147">Can contain</span></span>

|<span data-ttu-id="9c54b-148">**Elemento**</span><span class="sxs-lookup"><span data-stu-id="9c54b-148">**Element**</span></span>|<span data-ttu-id="9c54b-149">**Content**</span><span class="sxs-lookup"><span data-stu-id="9c54b-149">**Content**</span></span>|<span data-ttu-id="9c54b-150">**Email**</span><span class="sxs-lookup"><span data-stu-id="9c54b-150">**Mail**</span></span>|<span data-ttu-id="9c54b-151">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="9c54b-151">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="9c54b-152">AlternateId</span><span class="sxs-lookup"><span data-stu-id="9c54b-152">AlternateId</span></span>](alternateid.md)|<span data-ttu-id="9c54b-153">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-153">x</span></span>|<span data-ttu-id="9c54b-154">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-154">x</span></span>|<span data-ttu-id="9c54b-155">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-155">x</span></span>|
|[<span data-ttu-id="9c54b-156">IconUrl</span><span class="sxs-lookup"><span data-stu-id="9c54b-156">IconUrl</span></span>](iconurl.md)|<span data-ttu-id="9c54b-157">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-157">x</span></span>|<span data-ttu-id="9c54b-158">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-158">x</span></span>|<span data-ttu-id="9c54b-159">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-159">x</span></span>|
|[<span data-ttu-id="9c54b-160">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="9c54b-160">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|<span data-ttu-id="9c54b-161">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-161">x</span></span>|<span data-ttu-id="9c54b-162">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-162">x</span></span>|<span data-ttu-id="9c54b-163">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-163">x</span></span>|
|[<span data-ttu-id="9c54b-164">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="9c54b-164">SupportUrl</span></span>](supporturl.md)|<span data-ttu-id="9c54b-165">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-165">x</span></span>|<span data-ttu-id="9c54b-166">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-166">x</span></span>|<span data-ttu-id="9c54b-167">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-167">x</span></span>|
|[<span data-ttu-id="9c54b-168">AppDomains</span><span class="sxs-lookup"><span data-stu-id="9c54b-168">AppDomains</span></span>](appdomains.md)|<span data-ttu-id="9c54b-169">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-169">x</span></span>|<span data-ttu-id="9c54b-170">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-170">x</span></span>|<span data-ttu-id="9c54b-171">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-171">x</span></span>|
|[<span data-ttu-id="9c54b-172">Hosts</span><span class="sxs-lookup"><span data-stu-id="9c54b-172">Hosts</span></span>](hosts.md)|<span data-ttu-id="9c54b-173">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-173">x</span></span>|<span data-ttu-id="9c54b-174">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-174">x</span></span>|<span data-ttu-id="9c54b-175">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-175">x</span></span>|
|[<span data-ttu-id="9c54b-176">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9c54b-176">Requirements</span></span>](requirements.md)|<span data-ttu-id="9c54b-177">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-177">x</span></span>|<span data-ttu-id="9c54b-178">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-178">x</span></span>|<span data-ttu-id="9c54b-179">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-179">x</span></span>|
|[<span data-ttu-id="9c54b-180">AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="9c54b-180">AllowSnapshot</span></span>](allowsnapshot.md)|<span data-ttu-id="9c54b-181">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-181">x</span></span>|||
|[<span data-ttu-id="9c54b-182">Permissões</span><span class="sxs-lookup"><span data-stu-id="9c54b-182">Permissions</span></span>](permissions.md)||<span data-ttu-id="9c54b-183">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-183">x</span></span>||
|[<span data-ttu-id="9c54b-184">DisableEntityHighlighting</span><span class="sxs-lookup"><span data-stu-id="9c54b-184">DisableEntityHighlighting</span></span>](disableentityhighlighting.md)||<span data-ttu-id="9c54b-185">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-185">x</span></span>||
|[<span data-ttu-id="9c54b-186">Dictionary</span><span class="sxs-lookup"><span data-stu-id="9c54b-186">Dictionary</span></span>](dictionary.md)|||<span data-ttu-id="9c54b-187">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-187">x</span></span>|
|[<span data-ttu-id="9c54b-188">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="9c54b-188">VersionOverrides</span></span>](versionoverrides.md)|<span data-ttu-id="9c54b-189">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-189">x</span></span>|<span data-ttu-id="9c54b-190">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-190">x</span></span>|<span data-ttu-id="9c54b-191">x</span><span class="sxs-lookup"><span data-stu-id="9c54b-191">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="9c54b-192">Atributos</span><span class="sxs-lookup"><span data-stu-id="9c54b-192">Attributes</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="9c54b-193">xmlns</span><span class="sxs-lookup"><span data-stu-id="9c54b-193">xmlns</span></span>|<span data-ttu-id="9c54b-p101">Define o namespace do manifesto do Suplemento do Office e o esquema da versão. Esse atributo deve ser sempre definido como `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span><span class="sxs-lookup"><span data-stu-id="9c54b-p101">Defines the Office Add-in manifest namespace and schema version. This attribute should always be set to  `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span></span>|
|<span data-ttu-id="9c54b-196">xmlns: xsi</span><span class="sxs-lookup"><span data-stu-id="9c54b-196">xmlns:xsi</span></span>|<span data-ttu-id="9c54b-p102">Define a instância XMLSchema. Esse atributo deve ser sempre definido como `"http://www.w3.org/2001/XMLSchema-instance"`</span><span class="sxs-lookup"><span data-stu-id="9c54b-p102">Defines the XMLSchema instance. This attribute should always be set to  `"http://www.w3.org/2001/XMLSchema-instance"`</span></span>|
|<span data-ttu-id="9c54b-199">xsi:type</span><span class="sxs-lookup"><span data-stu-id="9c54b-199">xsi:type</span></span>|<span data-ttu-id="9c54b-p103">Define o tipo de Suplemento do Office. Esse atributo deve ser definido como um destes: `"ContentApp"`, `"MailApp"` ou `"TaskPaneApp"`</span><span class="sxs-lookup"><span data-stu-id="9c54b-p103">Defines the kind of Office Add-in. This attribute should be set to one of:  `"ContentApp"`,  `"MailApp"`, or  `"TaskPaneApp"`</span></span>|
