---
title: Elemento OfficeApp no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 42b6fe2e1c33322b90016d5e7ceec7b1bfe5b72d
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433162"
---
# <a name="officeapp-element"></a><span data-ttu-id="91535-102">Elemento OfficeApp</span><span class="sxs-lookup"><span data-stu-id="91535-102">OfficeApp element</span></span>

<span data-ttu-id="91535-103">O elemento raiz no manifesto de um Suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="91535-103">The root element in the manifest of an Office Add-in.</span></span>

<span data-ttu-id="91535-104">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="91535-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="91535-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="91535-105">Syntax</span></span>

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="91535-106">Contido em</span><span class="sxs-lookup"><span data-stu-id="91535-106">Contained in</span></span>

 <span data-ttu-id="91535-107">_none_</span><span class="sxs-lookup"><span data-stu-id="91535-107">_none_</span></span>

## <a name="must-contain"></a><span data-ttu-id="91535-108">Deve conter</span><span class="sxs-lookup"><span data-stu-id="91535-108">Must contain</span></span>

|<span data-ttu-id="91535-109">**Element**</span><span class="sxs-lookup"><span data-stu-id="91535-109">**Element**</span></span>|<span data-ttu-id="91535-110">**Conteúdo**</span><span class="sxs-lookup"><span data-stu-id="91535-110">**Content**</span></span>|<span data-ttu-id="91535-111">**Email**</span><span class="sxs-lookup"><span data-stu-id="91535-111">**Mail**</span></span>|<span data-ttu-id="91535-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="91535-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="91535-113">ID</span><span class="sxs-lookup"><span data-stu-id="91535-113">Id</span></span>](id.md)|<span data-ttu-id="91535-114">x</span><span class="sxs-lookup"><span data-stu-id="91535-114">x</span></span>|<span data-ttu-id="91535-115">x</span><span class="sxs-lookup"><span data-stu-id="91535-115">x</span></span>|<span data-ttu-id="91535-116">x</span><span class="sxs-lookup"><span data-stu-id="91535-116">x</span></span>|
|[<span data-ttu-id="91535-117">Versão</span><span class="sxs-lookup"><span data-stu-id="91535-117">Version</span></span>](version.md)|<span data-ttu-id="91535-118">x</span><span class="sxs-lookup"><span data-stu-id="91535-118">x</span></span>|<span data-ttu-id="91535-119">x</span><span class="sxs-lookup"><span data-stu-id="91535-119">x</span></span>|<span data-ttu-id="91535-120">x</span><span class="sxs-lookup"><span data-stu-id="91535-120">x</span></span>|
|[<span data-ttu-id="91535-121">ProviderName</span><span class="sxs-lookup"><span data-stu-id="91535-121">ProviderName</span></span>](providername.md)|<span data-ttu-id="91535-122">x</span><span class="sxs-lookup"><span data-stu-id="91535-122">x</span></span>|<span data-ttu-id="91535-123">x</span><span class="sxs-lookup"><span data-stu-id="91535-123">x</span></span>|<span data-ttu-id="91535-124">x</span><span class="sxs-lookup"><span data-stu-id="91535-124">x</span></span>|
|[<span data-ttu-id="91535-125">DefaultLocale</span><span class="sxs-lookup"><span data-stu-id="91535-125">DefaultLocale</span></span>](defaultlocale.md)|<span data-ttu-id="91535-126">x</span><span class="sxs-lookup"><span data-stu-id="91535-126">x</span></span>|<span data-ttu-id="91535-127">x</span><span class="sxs-lookup"><span data-stu-id="91535-127">x</span></span>|<span data-ttu-id="91535-128">x</span><span class="sxs-lookup"><span data-stu-id="91535-128">x</span></span>|
|[<span data-ttu-id="91535-129">DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="91535-129">DefaultSettings</span></span>](defaultsettings.md)|<span data-ttu-id="91535-130">x</span><span class="sxs-lookup"><span data-stu-id="91535-130">x</span></span>||<span data-ttu-id="91535-131">x</span><span class="sxs-lookup"><span data-stu-id="91535-131">x</span></span>|
|[<span data-ttu-id="91535-132">DisplayName</span><span class="sxs-lookup"><span data-stu-id="91535-132">DisplayName</span></span>](displayname.md)|<span data-ttu-id="91535-133">x</span><span class="sxs-lookup"><span data-stu-id="91535-133">x</span></span>|<span data-ttu-id="91535-134">x</span><span class="sxs-lookup"><span data-stu-id="91535-134">x</span></span>|<span data-ttu-id="91535-135">x</span><span class="sxs-lookup"><span data-stu-id="91535-135">x</span></span>|
|[<span data-ttu-id="91535-136">Descrição</span><span class="sxs-lookup"><span data-stu-id="91535-136">Description</span></span>](description.md)|<span data-ttu-id="91535-137">x</span><span class="sxs-lookup"><span data-stu-id="91535-137">x</span></span>|<span data-ttu-id="91535-138">x</span><span class="sxs-lookup"><span data-stu-id="91535-138">x</span></span>|<span data-ttu-id="91535-139">x</span><span class="sxs-lookup"><span data-stu-id="91535-139">x</span></span>|
|[<span data-ttu-id="91535-140">FormSettings</span><span class="sxs-lookup"><span data-stu-id="91535-140">FormSettings</span></span>](formsettings.md)||<span data-ttu-id="91535-141">x</span><span class="sxs-lookup"><span data-stu-id="91535-141">x</span></span>||
|[<span data-ttu-id="91535-142">Permissões</span><span class="sxs-lookup"><span data-stu-id="91535-142">Permissions</span></span>](permissions.md)|<span data-ttu-id="91535-143">x</span><span class="sxs-lookup"><span data-stu-id="91535-143">x</span></span>||<span data-ttu-id="91535-144">x</span><span class="sxs-lookup"><span data-stu-id="91535-144">x</span></span>|
|[<span data-ttu-id="91535-145">Rule</span><span class="sxs-lookup"><span data-stu-id="91535-145">Rule</span></span>](rule.md)||<span data-ttu-id="91535-146">x</span><span class="sxs-lookup"><span data-stu-id="91535-146">x</span></span>||

## <a name="can-contain"></a><span data-ttu-id="91535-147">Pode conter</span><span class="sxs-lookup"><span data-stu-id="91535-147">Can contain</span></span>

|<span data-ttu-id="91535-148">**Element**</span><span class="sxs-lookup"><span data-stu-id="91535-148">**Element**</span></span>|<span data-ttu-id="91535-149">**Conteúdo**</span><span class="sxs-lookup"><span data-stu-id="91535-149">**Content**</span></span>|<span data-ttu-id="91535-150">**Email**</span><span class="sxs-lookup"><span data-stu-id="91535-150">**Mail**</span></span>|<span data-ttu-id="91535-151">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="91535-151">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="91535-152">AlternateId</span><span class="sxs-lookup"><span data-stu-id="91535-152">AlternateId</span></span>](alternateid.md)|<span data-ttu-id="91535-153">x</span><span class="sxs-lookup"><span data-stu-id="91535-153">x</span></span>|<span data-ttu-id="91535-154">x</span><span class="sxs-lookup"><span data-stu-id="91535-154">x</span></span>|<span data-ttu-id="91535-155">x</span><span class="sxs-lookup"><span data-stu-id="91535-155">x</span></span>|
|[<span data-ttu-id="91535-156">IconUrl</span><span class="sxs-lookup"><span data-stu-id="91535-156">IconUrl</span></span>](iconurl.md)|<span data-ttu-id="91535-157">x</span><span class="sxs-lookup"><span data-stu-id="91535-157">x</span></span>|<span data-ttu-id="91535-158">x</span><span class="sxs-lookup"><span data-stu-id="91535-158">x</span></span>|<span data-ttu-id="91535-159">x</span><span class="sxs-lookup"><span data-stu-id="91535-159">x</span></span>|
|[<span data-ttu-id="91535-160">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="91535-160">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|<span data-ttu-id="91535-161">x</span><span class="sxs-lookup"><span data-stu-id="91535-161">x</span></span>|<span data-ttu-id="91535-162">x</span><span class="sxs-lookup"><span data-stu-id="91535-162">x</span></span>|<span data-ttu-id="91535-163">x</span><span class="sxs-lookup"><span data-stu-id="91535-163">x</span></span>|
|[<span data-ttu-id="91535-164">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="91535-164">SupportUrl</span></span>](supporturl.md)|<span data-ttu-id="91535-165">x</span><span class="sxs-lookup"><span data-stu-id="91535-165">x</span></span>|<span data-ttu-id="91535-166">x</span><span class="sxs-lookup"><span data-stu-id="91535-166">x</span></span>|<span data-ttu-id="91535-167">x</span><span class="sxs-lookup"><span data-stu-id="91535-167">x</span></span>|
|[<span data-ttu-id="91535-168">AppDomains</span><span class="sxs-lookup"><span data-stu-id="91535-168">AppDomains</span></span>](appdomains.md)|<span data-ttu-id="91535-169">x</span><span class="sxs-lookup"><span data-stu-id="91535-169">x</span></span>|<span data-ttu-id="91535-170">x</span><span class="sxs-lookup"><span data-stu-id="91535-170">x</span></span>|<span data-ttu-id="91535-171">x</span><span class="sxs-lookup"><span data-stu-id="91535-171">x</span></span>|
|[<span data-ttu-id="91535-172">Hosts</span><span class="sxs-lookup"><span data-stu-id="91535-172">Hosts</span></span>](hosts.md)|<span data-ttu-id="91535-173">x</span><span class="sxs-lookup"><span data-stu-id="91535-173">x</span></span>|<span data-ttu-id="91535-174">x</span><span class="sxs-lookup"><span data-stu-id="91535-174">x</span></span>|<span data-ttu-id="91535-175">x</span><span class="sxs-lookup"><span data-stu-id="91535-175">x</span></span>|
|[<span data-ttu-id="91535-176">Requirements</span><span class="sxs-lookup"><span data-stu-id="91535-176">Requirements</span></span>](requirements.md)|<span data-ttu-id="91535-177">x</span><span class="sxs-lookup"><span data-stu-id="91535-177">x</span></span>|<span data-ttu-id="91535-178">x</span><span class="sxs-lookup"><span data-stu-id="91535-178">x</span></span>|<span data-ttu-id="91535-179">x</span><span class="sxs-lookup"><span data-stu-id="91535-179">x</span></span>|
|[<span data-ttu-id="91535-180">AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="91535-180">AllowSnapshot</span></span>](allowsnapshot.md)|<span data-ttu-id="91535-181">x</span><span class="sxs-lookup"><span data-stu-id="91535-181">x</span></span>|||
|[<span data-ttu-id="91535-182">Permissões</span><span class="sxs-lookup"><span data-stu-id="91535-182">Permissions</span></span>](permissions.md)||<span data-ttu-id="91535-183">x</span><span class="sxs-lookup"><span data-stu-id="91535-183">x</span></span>||
|[<span data-ttu-id="91535-184">DisableEntityHighlighting</span><span class="sxs-lookup"><span data-stu-id="91535-184">DisableEntityHighlighting</span></span>](disableentityhighlighting.md)||<span data-ttu-id="91535-185">x</span><span class="sxs-lookup"><span data-stu-id="91535-185">x</span></span>||
|[<span data-ttu-id="91535-186">Dictionary</span><span class="sxs-lookup"><span data-stu-id="91535-186">Dictionary</span></span>](dictionary.md)|||<span data-ttu-id="91535-187">x</span><span class="sxs-lookup"><span data-stu-id="91535-187">x</span></span>|
|[<span data-ttu-id="91535-188">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="91535-188">VersionOverrides</span></span>](versionoverrides.md)||<span data-ttu-id="91535-189">x</span><span class="sxs-lookup"><span data-stu-id="91535-189">x</span></span>||

## <a name="attributes"></a><span data-ttu-id="91535-190">Atributos</span><span class="sxs-lookup"><span data-stu-id="91535-190">Attributes</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="91535-191">xmlns</span><span class="sxs-lookup"><span data-stu-id="91535-191">xmlns</span></span>|<span data-ttu-id="91535-p101">Define o namespace do manifesto do Suplemento do Office e o esquema da versão. Esse atributo deve ser sempre definido como `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span><span class="sxs-lookup"><span data-stu-id="91535-p101">Defines the Office Add-in manifest namespace and schema version. This attribute should always be set to  `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span></span>|
|<span data-ttu-id="91535-194">xmlns: xsi</span><span class="sxs-lookup"><span data-stu-id="91535-194">xmlns:xsi</span></span>|<span data-ttu-id="91535-p102">Define a instância XMLSchema. Esse atributo deve ser sempre definido como `"http://www.w3.org/2001/XMLSchema-instance"`</span><span class="sxs-lookup"><span data-stu-id="91535-p102">Defines the XMLSchema instance. This attribute should always be set to  `"http://www.w3.org/2001/XMLSchema-instance"`</span></span>|
|<span data-ttu-id="91535-197">xsi:type</span><span class="sxs-lookup"><span data-stu-id="91535-197">xsi:type</span></span>|<span data-ttu-id="91535-p103">Define o tipo de Suplemento do Office. Esse atributo deve ser definido como um destes: `"ContentApp"`, `"MailApp"` ou `"TaskPaneApp"`</span><span class="sxs-lookup"><span data-stu-id="91535-p103">Defines the kind of Office Add-in. This attribute should be set to one of:  `"ContentApp"`,  `"MailApp"`, or  `"TaskPaneApp"`</span></span>|
