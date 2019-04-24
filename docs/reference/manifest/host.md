---
title: Elemento Host no arquivo de manifesto
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: f496e3e0c16f24d20e1d1db76208e61267235131
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450503"
---
# <a name="host-element"></a><span data-ttu-id="ac541-102">Elemento Host</span><span class="sxs-lookup"><span data-stu-id="ac541-102">Host element</span></span>

<span data-ttu-id="ac541-103">Especifica um tipo de aplicativo individual do Office em que o suplemento deve ser ativado.</span><span class="sxs-lookup"><span data-stu-id="ac541-103">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="ac541-104">A sintaxe do elemento **Host** varia de acordo com a definição do elemento, se dentro do [manifesto básico](#basic-manifest) ou dentro do nó [VersionOverrides](#versionoverrides-node).</span><span class="sxs-lookup"><span data-stu-id="ac541-104">The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="ac541-105">No entanto, a funcionalidade é a mesma.</span><span class="sxs-lookup"><span data-stu-id="ac541-105">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="ac541-106">Manifesto básico</span><span class="sxs-lookup"><span data-stu-id="ac541-106">Basic manifest</span></span>

<span data-ttu-id="ac541-107">Quando definido no manifesto básico (em [OfficeApp](officeapp.md)), o tipo de host é determinado pelo atributo `Name`.</span><span class="sxs-lookup"><span data-stu-id="ac541-107">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>   

### <a name="attributes"></a><span data-ttu-id="ac541-108">Atributos</span><span class="sxs-lookup"><span data-stu-id="ac541-108">Attributes</span></span>

| <span data-ttu-id="ac541-109">Atributo</span><span class="sxs-lookup"><span data-stu-id="ac541-109">Attribute</span></span>     | <span data-ttu-id="ac541-110">Tipo</span><span class="sxs-lookup"><span data-stu-id="ac541-110">Type</span></span>   | <span data-ttu-id="ac541-111">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="ac541-111">Required</span></span> | <span data-ttu-id="ac541-112">Descrição</span><span class="sxs-lookup"><span data-stu-id="ac541-112">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="ac541-113">Nome</span><span class="sxs-lookup"><span data-stu-id="ac541-113">Name</span></span>](#name) | <span data-ttu-id="ac541-114">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ac541-114">string</span></span> | <span data-ttu-id="ac541-115">obrigatório</span><span class="sxs-lookup"><span data-stu-id="ac541-115">required</span></span> | <span data-ttu-id="ac541-116">O nome do tipo de aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="ac541-116">The name of the type of Office host application.</span></span> |

### <a name="name"></a><span data-ttu-id="ac541-117">Name</span><span class="sxs-lookup"><span data-stu-id="ac541-117">Name</span></span>
<span data-ttu-id="ac541-p102">Especifica o tipo de Host destinado por esse suplemento. O valor deve ser uma das seguintes opções:</span><span class="sxs-lookup"><span data-stu-id="ac541-p102">Specifies the Host type targeted by this add-in. The value must be one of the following:</span></span>

- <span data-ttu-id="ac541-120">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="ac541-120">`Document` (Word)</span></span>
- <span data-ttu-id="ac541-121">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="ac541-121">`Database` (Access)</span></span>
- <span data-ttu-id="ac541-122">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="ac541-122">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="ac541-123">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="ac541-123">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="ac541-124">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="ac541-124">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="ac541-125">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="ac541-125">`Project` (Project)</span></span>
- <span data-ttu-id="ac541-126">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="ac541-126">`Workbook` (Excel)</span></span>

### <a name="example"></a><span data-ttu-id="ac541-127">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ac541-127">Example</span></span>
```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="ac541-128">Nó VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="ac541-128">VersionOverrides node</span></span>
<span data-ttu-id="ac541-129">Quando definido em [VersionOverrides](versionoverrides.md), o tipo de host é determinado pelo atributo `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="ac541-129">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span> 

### <a name="attributes"></a><span data-ttu-id="ac541-130">Atributos</span><span class="sxs-lookup"><span data-stu-id="ac541-130">Attributes</span></span>

|  <span data-ttu-id="ac541-131">Atributo</span><span class="sxs-lookup"><span data-stu-id="ac541-131">Attribute</span></span>  |  <span data-ttu-id="ac541-132">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="ac541-132">Required</span></span>  |  <span data-ttu-id="ac541-133">Descrição</span><span class="sxs-lookup"><span data-stu-id="ac541-133">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="ac541-134">xsi:type</span><span class="sxs-lookup"><span data-stu-id="ac541-134">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="ac541-135">Sim</span><span class="sxs-lookup"><span data-stu-id="ac541-135">Yes</span></span>  | <span data-ttu-id="ac541-136">Descreve o host do Office a que essas configurações se aplicam.</span><span class="sxs-lookup"><span data-stu-id="ac541-136">Describes the Office host where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="ac541-137">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="ac541-137">Child elements</span></span>

|  <span data-ttu-id="ac541-138">Elemento</span><span class="sxs-lookup"><span data-stu-id="ac541-138">Element</span></span> |  <span data-ttu-id="ac541-139">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="ac541-139">Required</span></span>  |  <span data-ttu-id="ac541-140">Descrição</span><span class="sxs-lookup"><span data-stu-id="ac541-140">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="ac541-141">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="ac541-141">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="ac541-142">Sim</span><span class="sxs-lookup"><span data-stu-id="ac541-142">Yes</span></span>   |  <span data-ttu-id="ac541-143">Define as configurações do fator forma da área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="ac541-143">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="ac541-144">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="ac541-144">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="ac541-145">Não</span><span class="sxs-lookup"><span data-stu-id="ac541-145">No</span></span>   |  <span data-ttu-id="ac541-p103">Define as configurações do fator forma móvel. **Observação:** esse elemento só tem suporte no Outlook para iOS.</span><span class="sxs-lookup"><span data-stu-id="ac541-p103">Defines the settings for the mobile form factor. **Note:** this element is only supported in Outlook for iOS.</span></span> |
|  [<span data-ttu-id="ac541-148">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="ac541-148">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="ac541-149">Não</span><span class="sxs-lookup"><span data-stu-id="ac541-149">No</span></span>   |  <span data-ttu-id="ac541-150">Define as configurações de todos os fatores forma.</span><span class="sxs-lookup"><span data-stu-id="ac541-150">Defines the settings for all form factors.</span></span> <span data-ttu-id="ac541-151">Usado somente pelas funções personalizadas no Excel.</span><span class="sxs-lookup"><span data-stu-id="ac541-151">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="ac541-152">xsi:type</span><span class="sxs-lookup"><span data-stu-id="ac541-152">xsi:type</span></span>

<span data-ttu-id="ac541-153">Controla a qual host do Office (Word, Excel, PowerPoint, Outlook, OneNote) as configurações contidas se aplicam.</span><span class="sxs-lookup"><span data-stu-id="ac541-153">Controls which Office host (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="ac541-154">O valor deve ser uma das seguintes opções:</span><span class="sxs-lookup"><span data-stu-id="ac541-154">The value must be one of the following:</span></span>

- <span data-ttu-id="ac541-155">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="ac541-155">`Document` (Word)</span></span>
- <span data-ttu-id="ac541-156">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="ac541-156">`MailHost` (Outlook)</span></span>    
- <span data-ttu-id="ac541-157">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="ac541-157">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="ac541-158">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="ac541-158">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="ac541-159">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="ac541-159">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="ac541-160">Exemplo de host</span><span class="sxs-lookup"><span data-stu-id="ac541-160">Host example</span></span> 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
