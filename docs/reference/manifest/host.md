---
title: Elemento Host no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 37b772261ad82b4f899e73314a08ffd1dd03b442
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432806"
---
# <a name="host-element"></a><span data-ttu-id="4562f-102">Elemento Host</span><span class="sxs-lookup"><span data-stu-id="4562f-102">Host element</span></span>

<span data-ttu-id="4562f-103">Especifica um tipo de aplicativo individual do Office em que o suplemento deve ser ativado.</span><span class="sxs-lookup"><span data-stu-id="4562f-103">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="4562f-104">A sintaxe do elemento **Host** varia de acordo com a definição do elemento, se dentro do [manifesto básico](#basic-manifest) ou dentro do nó [VersionOverrides](#versionoverrides-node).</span><span class="sxs-lookup"><span data-stu-id="4562f-104">Important: The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="4562f-105">No entanto, a funcionalidade é a mesma.</span><span class="sxs-lookup"><span data-stu-id="4562f-105">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="4562f-106">Manifesto básico</span><span class="sxs-lookup"><span data-stu-id="4562f-106">Basic manifest</span></span>

<span data-ttu-id="4562f-107">Quando definido no manifesto básico (em [OfficeApp](officeapp.md)), o tipo de host é determinado pelo atributo `Name`.</span><span class="sxs-lookup"><span data-stu-id="4562f-107">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>   

### <a name="attributes"></a><span data-ttu-id="4562f-108">Atributos</span><span class="sxs-lookup"><span data-stu-id="4562f-108">Attributes</span></span>

| <span data-ttu-id="4562f-109">Atributo</span><span class="sxs-lookup"><span data-stu-id="4562f-109">Attribute</span></span>     | <span data-ttu-id="4562f-110">Tipo</span><span class="sxs-lookup"><span data-stu-id="4562f-110">Type</span></span>   | <span data-ttu-id="4562f-111">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="4562f-111">Required</span></span> | <span data-ttu-id="4562f-112">Descrição</span><span class="sxs-lookup"><span data-stu-id="4562f-112">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="4562f-113">Nome</span><span class="sxs-lookup"><span data-stu-id="4562f-113">Name</span></span>](#name) | <span data-ttu-id="4562f-114">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4562f-114">string</span></span> | <span data-ttu-id="4562f-115">obrigatório</span><span class="sxs-lookup"><span data-stu-id="4562f-115">required</span></span> | <span data-ttu-id="4562f-116">O nome do tipo de aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="4562f-116">The name of the type of Office host application.</span></span> |

### <a name="name"></a><span data-ttu-id="4562f-117">Name</span><span class="sxs-lookup"><span data-stu-id="4562f-117">Name</span></span>
<span data-ttu-id="4562f-p102">Especifica o tipo de Host destinado por esse suplemento. O valor deve ser uma das seguintes opções:</span><span class="sxs-lookup"><span data-stu-id="4562f-p102">Specifies the Host type targeted by this add-in. The value must be one of the following:</span></span>

- <span data-ttu-id="4562f-120">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="4562f-120">`Document` (Word)</span></span>
- <span data-ttu-id="4562f-121">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="4562f-121">`Database` (Access)</span></span>
- <span data-ttu-id="4562f-122">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="4562f-122">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="4562f-123">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="4562f-123">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="4562f-124">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="4562f-124">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="4562f-125">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="4562f-125">`Project` (Project)</span></span>
- <span data-ttu-id="4562f-126">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="4562f-126">`Workbook` (Excel)</span></span>

### <a name="example"></a><span data-ttu-id="4562f-127">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4562f-127">Example</span></span>
```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="4562f-128">Nó VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="4562f-128">VersionOverrides node</span></span>
<span data-ttu-id="4562f-129">Quando definido em [VersionOverrides](versionoverrides.md), o tipo de host é determinado pelo atributo `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="4562f-129">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span> 

### <a name="attributes"></a><span data-ttu-id="4562f-130">Atributos</span><span class="sxs-lookup"><span data-stu-id="4562f-130">Attributes</span></span>

|  <span data-ttu-id="4562f-131">Atributo</span><span class="sxs-lookup"><span data-stu-id="4562f-131">Attribute</span></span>  |  <span data-ttu-id="4562f-132">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="4562f-132">Required</span></span>  |  <span data-ttu-id="4562f-133">Descrição</span><span class="sxs-lookup"><span data-stu-id="4562f-133">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="4562f-134">xsi:type</span><span class="sxs-lookup"><span data-stu-id="4562f-134">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="4562f-135">Sim</span><span class="sxs-lookup"><span data-stu-id="4562f-135">Yes</span></span>  | <span data-ttu-id="4562f-136">Descreve o host do Office a que essas configurações se aplicam.</span><span class="sxs-lookup"><span data-stu-id="4562f-136">Describes the Office host where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="4562f-137">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="4562f-137">Child elements</span></span>

|  <span data-ttu-id="4562f-138">Elemento</span><span class="sxs-lookup"><span data-stu-id="4562f-138">Element</span></span> |  <span data-ttu-id="4562f-139">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="4562f-139">Required</span></span>  |  <span data-ttu-id="4562f-140">Descrição</span><span class="sxs-lookup"><span data-stu-id="4562f-140">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="4562f-141">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="4562f-141">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="4562f-142">Sim</span><span class="sxs-lookup"><span data-stu-id="4562f-142">Yes</span></span>   |  <span data-ttu-id="4562f-143">Define as configurações do fator forma da área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="4562f-143">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="4562f-144">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="4562f-144">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="4562f-145">Não</span><span class="sxs-lookup"><span data-stu-id="4562f-145">No</span></span>   |  <span data-ttu-id="4562f-p103">Define as configurações do fator forma móvel. **Observação:** esse elemento só tem suporte no Outlook para iOS.</span><span class="sxs-lookup"><span data-stu-id="4562f-p103">Defines the settings for the mobile form factor. **Note:** this element is only supported in Outlook for iOS.</span></span> |
|  [<span data-ttu-id="4562f-148">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="4562f-148">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="4562f-149">Não</span><span class="sxs-lookup"><span data-stu-id="4562f-149">No</span></span>   |  <span data-ttu-id="4562f-150">Define as configurações de todos os fatores forma.</span><span class="sxs-lookup"><span data-stu-id="4562f-150">Defines the settings for all form factors.</span></span> <span data-ttu-id="4562f-151">Usado somente pelas funções personalizadas no Excel.</span><span class="sxs-lookup"><span data-stu-id="4562f-151">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="4562f-152">xsi:type</span><span class="sxs-lookup"><span data-stu-id="4562f-152">xsi:type</span></span>

<span data-ttu-id="4562f-153">Controla a qual host do Office (Word, Excel, PowerPoint, Outlook, OneNote) as configurações contidas se aplicam.</span><span class="sxs-lookup"><span data-stu-id="4562f-153">Controls which Office host (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="4562f-154">O valor deve ser uma das seguintes opções:</span><span class="sxs-lookup"><span data-stu-id="4562f-154">The value must be one of the following:</span></span>

- <span data-ttu-id="4562f-155">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="4562f-155">`Document` (Word)</span></span>
- <span data-ttu-id="4562f-156">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="4562f-156">`MailHost` (Outlook)</span></span>    
- <span data-ttu-id="4562f-157">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="4562f-157">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="4562f-158">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="4562f-158">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="4562f-159">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="4562f-159">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="4562f-160">Exemplo de host</span><span class="sxs-lookup"><span data-stu-id="4562f-160">Host example</span></span> 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
