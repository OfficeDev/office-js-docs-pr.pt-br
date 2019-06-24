---
title: Elemento Host no arquivo de manifesto
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: debb4d59f75ce974ffb21d853c6b65a579c4e685
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127566"
---
# <a name="host-element"></a><span data-ttu-id="d10ef-102">Elemento Host</span><span class="sxs-lookup"><span data-stu-id="d10ef-102">Host element</span></span>

<span data-ttu-id="d10ef-103">Especifica um tipo de aplicativo individual do Office em que o suplemento deve ser ativado.</span><span class="sxs-lookup"><span data-stu-id="d10ef-103">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="d10ef-104">A sintaxe do elemento **Host** varia de acordo com a definição do elemento, se dentro do [manifesto básico](#basic-manifest) ou dentro do nó [VersionOverrides](#versionoverrides-node).</span><span class="sxs-lookup"><span data-stu-id="d10ef-104">The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="d10ef-105">No entanto, a funcionalidade é a mesma.</span><span class="sxs-lookup"><span data-stu-id="d10ef-105">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="d10ef-106">Manifesto básico</span><span class="sxs-lookup"><span data-stu-id="d10ef-106">Basic manifest</span></span>

<span data-ttu-id="d10ef-107">Quando definido no manifesto básico (em [OfficeApp](officeapp.md)), o tipo de host é determinado pelo atributo `Name`.</span><span class="sxs-lookup"><span data-stu-id="d10ef-107">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="d10ef-108">Atributos</span><span class="sxs-lookup"><span data-stu-id="d10ef-108">Attributes</span></span>

| <span data-ttu-id="d10ef-109">Atributo</span><span class="sxs-lookup"><span data-stu-id="d10ef-109">Attribute</span></span>     | <span data-ttu-id="d10ef-110">Tipo</span><span class="sxs-lookup"><span data-stu-id="d10ef-110">Type</span></span>   | <span data-ttu-id="d10ef-111">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="d10ef-111">Required</span></span> | <span data-ttu-id="d10ef-112">Descrição</span><span class="sxs-lookup"><span data-stu-id="d10ef-112">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="d10ef-113">Nome</span><span class="sxs-lookup"><span data-stu-id="d10ef-113">Name</span></span>](#name) | <span data-ttu-id="d10ef-114">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="d10ef-114">string</span></span> | <span data-ttu-id="d10ef-115">obrigatório</span><span class="sxs-lookup"><span data-stu-id="d10ef-115">required</span></span> | <span data-ttu-id="d10ef-116">O nome do tipo de aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="d10ef-116">The name of the type of Office host application.</span></span> |

### <a name="name"></a><span data-ttu-id="d10ef-117">Name</span><span class="sxs-lookup"><span data-stu-id="d10ef-117">Name</span></span>
<span data-ttu-id="d10ef-p102">Especifica o tipo de Host destinado por esse suplemento. O valor deve ser uma das seguintes opções:</span><span class="sxs-lookup"><span data-stu-id="d10ef-p102">Specifies the Host type targeted by this add-in. The value must be one of the following:</span></span>

- <span data-ttu-id="d10ef-120">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="d10ef-120">`Document` (Word)</span></span>
- <span data-ttu-id="d10ef-121">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="d10ef-121">`Database` (Access)</span></span>
- <span data-ttu-id="d10ef-122">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="d10ef-122">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="d10ef-123">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="d10ef-123">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="d10ef-124">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="d10ef-124">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="d10ef-125">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="d10ef-125">`Project` (Project)</span></span>
- <span data-ttu-id="d10ef-126">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="d10ef-126">`Workbook` (Excel)</span></span>

### <a name="example"></a><span data-ttu-id="d10ef-127">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d10ef-127">Example</span></span>
```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="d10ef-128">Nó VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="d10ef-128">VersionOverrides node</span></span>
<span data-ttu-id="d10ef-129">Quando definido em [VersionOverrides](versionoverrides.md), o tipo de host é determinado pelo atributo `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="d10ef-129">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span> 

### <a name="attributes"></a><span data-ttu-id="d10ef-130">Atributos</span><span class="sxs-lookup"><span data-stu-id="d10ef-130">Attributes</span></span>

|  <span data-ttu-id="d10ef-131">Atributo</span><span class="sxs-lookup"><span data-stu-id="d10ef-131">Attribute</span></span>  |  <span data-ttu-id="d10ef-132">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="d10ef-132">Required</span></span>  |  <span data-ttu-id="d10ef-133">Descrição</span><span class="sxs-lookup"><span data-stu-id="d10ef-133">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="d10ef-134">xsi:type</span><span class="sxs-lookup"><span data-stu-id="d10ef-134">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="d10ef-135">Sim</span><span class="sxs-lookup"><span data-stu-id="d10ef-135">Yes</span></span>  | <span data-ttu-id="d10ef-136">Descreve o host do Office a que essas configurações se aplicam.</span><span class="sxs-lookup"><span data-stu-id="d10ef-136">Describes the Office host where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="d10ef-137">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="d10ef-137">Child elements</span></span>

|  <span data-ttu-id="d10ef-138">Elemento</span><span class="sxs-lookup"><span data-stu-id="d10ef-138">Element</span></span> |  <span data-ttu-id="d10ef-139">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="d10ef-139">Required</span></span>  |  <span data-ttu-id="d10ef-140">Descrição</span><span class="sxs-lookup"><span data-stu-id="d10ef-140">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="d10ef-141">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="d10ef-141">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="d10ef-142">Sim</span><span class="sxs-lookup"><span data-stu-id="d10ef-142">Yes</span></span>   |  <span data-ttu-id="d10ef-143">Define as configurações do fator forma da área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="d10ef-143">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="d10ef-144">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="d10ef-144">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="d10ef-145">Não</span><span class="sxs-lookup"><span data-stu-id="d10ef-145">No</span></span>   |  <span data-ttu-id="d10ef-146">Define as configurações do fator forma móvel.</span><span class="sxs-lookup"><span data-stu-id="d10ef-146">Defines the settings for the mobile form factor.</span></span> <span data-ttu-id="d10ef-147">**Observação:** Esse elemento só é suportado no Outlook no iOS.</span><span class="sxs-lookup"><span data-stu-id="d10ef-147">**Note:** This element is only supported in Outlook on iOS.</span></span> |
|  [<span data-ttu-id="d10ef-148">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="d10ef-148">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="d10ef-149">Não</span><span class="sxs-lookup"><span data-stu-id="d10ef-149">No</span></span>   |  <span data-ttu-id="d10ef-150">Define as configurações de todos os fatores forma.</span><span class="sxs-lookup"><span data-stu-id="d10ef-150">Defines the settings for all form factors.</span></span> <span data-ttu-id="d10ef-151">Usado somente pelas funções personalizadas no Excel.</span><span class="sxs-lookup"><span data-stu-id="d10ef-151">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="d10ef-152">xsi:type</span><span class="sxs-lookup"><span data-stu-id="d10ef-152">xsi:type</span></span>

<span data-ttu-id="d10ef-153">Controla a qual host do Office (Word, Excel, PowerPoint, Outlook, OneNote) as configurações contidas se aplicam.</span><span class="sxs-lookup"><span data-stu-id="d10ef-153">Controls which Office host (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="d10ef-154">O valor deve ser uma das seguintes opções:</span><span class="sxs-lookup"><span data-stu-id="d10ef-154">The value must be one of the following:</span></span>

- <span data-ttu-id="d10ef-155">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="d10ef-155">`Document` (Word)</span></span>
- <span data-ttu-id="d10ef-156">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="d10ef-156">`MailHost` (Outlook)</span></span>
- <span data-ttu-id="d10ef-157">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="d10ef-157">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="d10ef-158">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="d10ef-158">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="d10ef-159">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="d10ef-159">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="d10ef-160">Exemplo de host</span><span class="sxs-lookup"><span data-stu-id="d10ef-160">Host example</span></span> 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
