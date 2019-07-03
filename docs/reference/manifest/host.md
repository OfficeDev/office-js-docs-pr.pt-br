---
title: Elemento Host no arquivo de manifesto
description: ''
ms.date: 07/01/2019
localization_priority: Normal
ms.openlocfilehash: e7b557034f70b03ed57598b7ffb9f43878db7392
ms.sourcegitcommit: 90c2d8236c6b30d80ac2b13950028a208ef60973
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/02/2019
ms.locfileid: "35454892"
---
# <a name="host-element"></a><span data-ttu-id="3ccb4-102">Elemento Host</span><span class="sxs-lookup"><span data-stu-id="3ccb4-102">Host element</span></span>

<span data-ttu-id="3ccb4-103">Especifica um tipo de aplicativo individual do Office em que o suplemento deve ser ativado.</span><span class="sxs-lookup"><span data-stu-id="3ccb4-103">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="3ccb4-104">A sintaxe do elemento **Host** varia de acordo com a definição do elemento, se dentro do [manifesto básico](#basic-manifest) ou dentro do nó [VersionOverrides](#versionoverrides-node).</span><span class="sxs-lookup"><span data-stu-id="3ccb4-104">The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="3ccb4-105">No entanto, a funcionalidade é a mesma.</span><span class="sxs-lookup"><span data-stu-id="3ccb4-105">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="3ccb4-106">Manifesto básico</span><span class="sxs-lookup"><span data-stu-id="3ccb4-106">Basic manifest</span></span>

<span data-ttu-id="3ccb4-107">Quando definido no manifesto básico (em [OfficeApp](officeapp.md)), o tipo de host é determinado pelo atributo `Name`.</span><span class="sxs-lookup"><span data-stu-id="3ccb4-107">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="3ccb4-108">Atributos</span><span class="sxs-lookup"><span data-stu-id="3ccb4-108">Attributes</span></span>

| <span data-ttu-id="3ccb4-109">Atributo</span><span class="sxs-lookup"><span data-stu-id="3ccb4-109">Attribute</span></span>     | <span data-ttu-id="3ccb4-110">Tipo</span><span class="sxs-lookup"><span data-stu-id="3ccb4-110">Type</span></span>   | <span data-ttu-id="3ccb4-111">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="3ccb4-111">Required</span></span> | <span data-ttu-id="3ccb4-112">Descrição</span><span class="sxs-lookup"><span data-stu-id="3ccb4-112">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="3ccb4-113">Nome</span><span class="sxs-lookup"><span data-stu-id="3ccb4-113">Name</span></span>](#name) | <span data-ttu-id="3ccb4-114">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="3ccb4-114">string</span></span> | <span data-ttu-id="3ccb4-115">obrigatório</span><span class="sxs-lookup"><span data-stu-id="3ccb4-115">required</span></span> | <span data-ttu-id="3ccb4-116">O nome do tipo de aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="3ccb4-116">The name of the type of Office host application.</span></span> |

### <a name="name"></a><span data-ttu-id="3ccb4-117">Name</span><span class="sxs-lookup"><span data-stu-id="3ccb4-117">Name</span></span>

<span data-ttu-id="3ccb4-118">Especifica o tipo de Host destinado por esse suplemento.</span><span class="sxs-lookup"><span data-stu-id="3ccb4-118">Specifies the Host type targeted by this add-in.</span></span> <span data-ttu-id="3ccb4-119">O valor deve ser um dos seguintes.</span><span class="sxs-lookup"><span data-stu-id="3ccb4-119">The value must be one of the following.</span></span>

- <span data-ttu-id="3ccb4-120">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="3ccb4-120">`Document` (Word)</span></span>
- <span data-ttu-id="3ccb4-121">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="3ccb4-121">`Database` (Access)</span></span>
- <span data-ttu-id="3ccb4-122">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="3ccb4-122">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="3ccb4-123">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="3ccb4-123">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="3ccb4-124">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="3ccb4-124">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="3ccb4-125">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="3ccb4-125">`Project` (Project)</span></span>
- <span data-ttu-id="3ccb4-126">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="3ccb4-126">`Workbook` (Excel)</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3ccb4-127">Não recomendamos mais criar e usar aplicativos Web do Access e bancos de dados no SharePoint.</span><span class="sxs-lookup"><span data-stu-id="3ccb4-127">We no longer recommend that you create and use Access web apps and databases in SharePoint.</span></span> <span data-ttu-id="3ccb4-128">Como alternativa, use o [Microsoft PowerApps](https://powerapps.microsoft.com/) para criar soluções de negócios sem código para dispositivos móveis e Web.</span><span class="sxs-lookup"><span data-stu-id="3ccb4-128">As an alternative, we recommend that you use [Microsoft PowerApps](https://powerapps.microsoft.com/) to build no-code business solutions for web and mobile devices.</span></span>

### <a name="example"></a><span data-ttu-id="3ccb4-129">Exemplo</span><span class="sxs-lookup"><span data-stu-id="3ccb4-129">Example</span></span>

```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="3ccb4-130">Nó VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="3ccb4-130">VersionOverrides node</span></span>

<span data-ttu-id="3ccb4-131">Quando definido em [VersionOverrides](versionoverrides.md), o tipo de host é determinado pelo atributo `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="3ccb4-131">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span> 

### <a name="attributes"></a><span data-ttu-id="3ccb4-132">Atributos</span><span class="sxs-lookup"><span data-stu-id="3ccb4-132">Attributes</span></span>

|  <span data-ttu-id="3ccb4-133">Atributo</span><span class="sxs-lookup"><span data-stu-id="3ccb4-133">Attribute</span></span>  |  <span data-ttu-id="3ccb4-134">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="3ccb4-134">Required</span></span>  |  <span data-ttu-id="3ccb4-135">Descrição</span><span class="sxs-lookup"><span data-stu-id="3ccb4-135">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="3ccb4-136">xsi:type</span><span class="sxs-lookup"><span data-stu-id="3ccb4-136">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="3ccb4-137">Sim</span><span class="sxs-lookup"><span data-stu-id="3ccb4-137">Yes</span></span>  | <span data-ttu-id="3ccb4-138">Descreve o host do Office a que essas configurações se aplicam.</span><span class="sxs-lookup"><span data-stu-id="3ccb4-138">Describes the Office host where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="3ccb4-139">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="3ccb4-139">Child elements</span></span>

|  <span data-ttu-id="3ccb4-140">Elemento</span><span class="sxs-lookup"><span data-stu-id="3ccb4-140">Element</span></span> |  <span data-ttu-id="3ccb4-141">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="3ccb4-141">Required</span></span>  |  <span data-ttu-id="3ccb4-142">Descrição</span><span class="sxs-lookup"><span data-stu-id="3ccb4-142">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="3ccb4-143">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="3ccb4-143">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="3ccb4-144">Sim</span><span class="sxs-lookup"><span data-stu-id="3ccb4-144">Yes</span></span>   |  <span data-ttu-id="3ccb4-145">Define as configurações do fator forma da área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="3ccb4-145">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="3ccb4-146">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="3ccb4-146">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="3ccb4-147">Não</span><span class="sxs-lookup"><span data-stu-id="3ccb4-147">No</span></span>   |  <span data-ttu-id="3ccb4-148">Define as configurações do fator forma móvel.</span><span class="sxs-lookup"><span data-stu-id="3ccb4-148">Defines the settings for the mobile form factor.</span></span> <span data-ttu-id="3ccb4-149">**Observação:** Esse elemento só é suportado no Outlook no iOS.</span><span class="sxs-lookup"><span data-stu-id="3ccb4-149">**Note:** This element is only supported in Outlook on iOS.</span></span> |
|  [<span data-ttu-id="3ccb4-150">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="3ccb4-150">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="3ccb4-151">Não</span><span class="sxs-lookup"><span data-stu-id="3ccb4-151">No</span></span>   |  <span data-ttu-id="3ccb4-152">Define as configurações de todos os fatores forma.</span><span class="sxs-lookup"><span data-stu-id="3ccb4-152">Defines the settings for all form factors.</span></span> <span data-ttu-id="3ccb4-153">Usado somente pelas funções personalizadas no Excel.</span><span class="sxs-lookup"><span data-stu-id="3ccb4-153">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="3ccb4-154">xsi:type</span><span class="sxs-lookup"><span data-stu-id="3ccb4-154">xsi:type</span></span>

<span data-ttu-id="3ccb4-155">Controla a qual host do Office (Word, Excel, PowerPoint, Outlook, OneNote) as configurações contidas se aplicam.</span><span class="sxs-lookup"><span data-stu-id="3ccb4-155">Controls which Office host (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="3ccb4-156">O valor deve ser uma das seguintes opções:</span><span class="sxs-lookup"><span data-stu-id="3ccb4-156">The value must be one of the following:</span></span>

- <span data-ttu-id="3ccb4-157">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="3ccb4-157">`Document` (Word)</span></span>
- <span data-ttu-id="3ccb4-158">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="3ccb4-158">`MailHost` (Outlook)</span></span>
- <span data-ttu-id="3ccb4-159">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="3ccb4-159">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="3ccb4-160">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="3ccb4-160">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="3ccb4-161">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="3ccb4-161">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="3ccb4-162">Exemplo de host</span><span class="sxs-lookup"><span data-stu-id="3ccb4-162">Host example</span></span> 

```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
