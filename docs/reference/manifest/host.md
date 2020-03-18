---
title: Elemento Host no arquivo de manifesto
description: Especifica um tipo de aplicativo individual do Office em que o suplemento deve ser ativado.
ms.date: 11/05/2019
localization_priority: Normal
ms.openlocfilehash: b9f03e6d6b028ca6f4616ae81b8fd76601256793
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718129"
---
# <a name="host-element"></a><span data-ttu-id="745d9-103">Elemento Host</span><span class="sxs-lookup"><span data-stu-id="745d9-103">Host element</span></span>

<span data-ttu-id="745d9-104">Especifica um tipo de aplicativo individual do Office em que o suplemento deve ser ativado.</span><span class="sxs-lookup"><span data-stu-id="745d9-104">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="745d9-105">A sintaxe do elemento **Host** varia de acordo com a definição do elemento, se dentro do [manifesto básico](#basic-manifest) ou dentro do nó [VersionOverrides](#versionoverrides-node).</span><span class="sxs-lookup"><span data-stu-id="745d9-105">The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="745d9-106">No entanto, a funcionalidade é a mesma.</span><span class="sxs-lookup"><span data-stu-id="745d9-106">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="745d9-107">Manifesto básico</span><span class="sxs-lookup"><span data-stu-id="745d9-107">Basic manifest</span></span>

<span data-ttu-id="745d9-108">Quando definido no manifesto básico (em [OfficeApp](officeapp.md)), o tipo de host é determinado pelo atributo `Name`.</span><span class="sxs-lookup"><span data-stu-id="745d9-108">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="745d9-109">Atributos</span><span class="sxs-lookup"><span data-stu-id="745d9-109">Attributes</span></span>

| <span data-ttu-id="745d9-110">Atributo</span><span class="sxs-lookup"><span data-stu-id="745d9-110">Attribute</span></span>     | <span data-ttu-id="745d9-111">Tipo</span><span class="sxs-lookup"><span data-stu-id="745d9-111">Type</span></span>   | <span data-ttu-id="745d9-112">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="745d9-112">Required</span></span> | <span data-ttu-id="745d9-113">Descrição</span><span class="sxs-lookup"><span data-stu-id="745d9-113">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="745d9-114">Nome</span><span class="sxs-lookup"><span data-stu-id="745d9-114">Name</span></span>](#name) | <span data-ttu-id="745d9-115">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="745d9-115">string</span></span> | <span data-ttu-id="745d9-116">obrigatório</span><span class="sxs-lookup"><span data-stu-id="745d9-116">required</span></span> | <span data-ttu-id="745d9-117">O nome do tipo de aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="745d9-117">The name of the type of Office host application.</span></span> |

### <a name="name"></a><span data-ttu-id="745d9-118">Name</span><span class="sxs-lookup"><span data-stu-id="745d9-118">Name</span></span>

<span data-ttu-id="745d9-119">Especifica o tipo de Host destinado por esse suplemento.</span><span class="sxs-lookup"><span data-stu-id="745d9-119">Specifies the Host type targeted by this add-in.</span></span> <span data-ttu-id="745d9-120">O valor deve ser um dos seguintes.</span><span class="sxs-lookup"><span data-stu-id="745d9-120">The value must be one of the following.</span></span>

- <span data-ttu-id="745d9-121">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="745d9-121">`Document` (Word)</span></span>
- <span data-ttu-id="745d9-122">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="745d9-122">`Database` (Access)</span></span>
- <span data-ttu-id="745d9-123">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="745d9-123">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="745d9-124">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="745d9-124">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="745d9-125">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="745d9-125">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="745d9-126">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="745d9-126">`Project` (Project)</span></span>
- <span data-ttu-id="745d9-127">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="745d9-127">`Workbook` (Excel)</span></span>

> [!IMPORTANT]
> <span data-ttu-id="745d9-128">Não recomendamos mais criar e usar aplicativos Web do Access e bancos de dados no SharePoint.</span><span class="sxs-lookup"><span data-stu-id="745d9-128">We no longer recommend that you create and use Access web apps and databases in SharePoint.</span></span> <span data-ttu-id="745d9-129">Como alternativa, use o [Microsoft PowerApps](https://powerapps.microsoft.com/) para criar soluções de negócios sem código para dispositivos móveis e Web.</span><span class="sxs-lookup"><span data-stu-id="745d9-129">As an alternative, we recommend that you use [Microsoft PowerApps](https://powerapps.microsoft.com/) to build no-code business solutions for web and mobile devices.</span></span>

### <a name="example"></a><span data-ttu-id="745d9-130">Exemplo</span><span class="sxs-lookup"><span data-stu-id="745d9-130">Example</span></span>

```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="745d9-131">Nó VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="745d9-131">VersionOverrides node</span></span>

<span data-ttu-id="745d9-132">Quando definido em [VersionOverrides](versionoverrides.md), o tipo de host é determinado pelo atributo `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="745d9-132">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="745d9-133">Atributos</span><span class="sxs-lookup"><span data-stu-id="745d9-133">Attributes</span></span>

|  <span data-ttu-id="745d9-134">Atributo</span><span class="sxs-lookup"><span data-stu-id="745d9-134">Attribute</span></span>  |  <span data-ttu-id="745d9-135">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="745d9-135">Required</span></span>  |  <span data-ttu-id="745d9-136">Descrição</span><span class="sxs-lookup"><span data-stu-id="745d9-136">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="745d9-137">xsi:type</span><span class="sxs-lookup"><span data-stu-id="745d9-137">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="745d9-138">Sim</span><span class="sxs-lookup"><span data-stu-id="745d9-138">Yes</span></span>  | <span data-ttu-id="745d9-139">Descreve o host do Office a que essas configurações se aplicam.</span><span class="sxs-lookup"><span data-stu-id="745d9-139">Describes the Office host where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="745d9-140">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="745d9-140">Child elements</span></span>

|  <span data-ttu-id="745d9-141">Elemento</span><span class="sxs-lookup"><span data-stu-id="745d9-141">Element</span></span> |  <span data-ttu-id="745d9-142">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="745d9-142">Required</span></span>  |  <span data-ttu-id="745d9-143">Descrição</span><span class="sxs-lookup"><span data-stu-id="745d9-143">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="745d9-144">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="745d9-144">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="745d9-145">Sim</span><span class="sxs-lookup"><span data-stu-id="745d9-145">Yes</span></span>   |  <span data-ttu-id="745d9-146">Define as configurações do fator forma da área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="745d9-146">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="745d9-147">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="745d9-147">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="745d9-148">Não</span><span class="sxs-lookup"><span data-stu-id="745d9-148">No</span></span>   |  <span data-ttu-id="745d9-149">Define as configurações do fator forma móvel.</span><span class="sxs-lookup"><span data-stu-id="745d9-149">Defines the settings for the mobile form factor.</span></span> <span data-ttu-id="745d9-150">**Observação:** Esse elemento só é suportado no Outlook no iOS e no Android.</span><span class="sxs-lookup"><span data-stu-id="745d9-150">**Note:** This element is only supported in Outlook on iOS and Android.</span></span> |
|  [<span data-ttu-id="745d9-151">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="745d9-151">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="745d9-152">Não</span><span class="sxs-lookup"><span data-stu-id="745d9-152">No</span></span>   |  <span data-ttu-id="745d9-153">Define as configurações de todos os fatores forma.</span><span class="sxs-lookup"><span data-stu-id="745d9-153">Defines the settings for all form factors.</span></span> <span data-ttu-id="745d9-154">Usado somente pelas funções personalizadas no Excel.</span><span class="sxs-lookup"><span data-stu-id="745d9-154">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="745d9-155">xsi:type</span><span class="sxs-lookup"><span data-stu-id="745d9-155">xsi:type</span></span>

<span data-ttu-id="745d9-156">Controla a qual host do Office (Word, Excel, PowerPoint, Outlook, OneNote) as configurações contidas se aplicam.</span><span class="sxs-lookup"><span data-stu-id="745d9-156">Controls which Office host (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="745d9-157">O valor deve ser uma das seguintes opções:</span><span class="sxs-lookup"><span data-stu-id="745d9-157">The value must be one of the following:</span></span>

- <span data-ttu-id="745d9-158">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="745d9-158">`Document` (Word)</span></span>
- <span data-ttu-id="745d9-159">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="745d9-159">`MailHost` (Outlook)</span></span>
- <span data-ttu-id="745d9-160">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="745d9-160">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="745d9-161">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="745d9-161">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="745d9-162">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="745d9-162">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="745d9-163">Exemplo de host</span><span class="sxs-lookup"><span data-stu-id="745d9-163">Host example</span></span>

```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
