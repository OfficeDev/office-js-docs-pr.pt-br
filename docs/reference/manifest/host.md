---
title: Elemento Host no arquivo de manifesto
description: Especifica um tipo de aplicativo individual do Office em que o suplemento deve ser ativado.
ms.date: 11/05/2019
localization_priority: Normal
ms.openlocfilehash: 5db9df97c4ba558d54756b983a26cb7b71e049d5
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611810"
---
# <a name="host-element"></a><span data-ttu-id="441b3-103">Elemento Host</span><span class="sxs-lookup"><span data-stu-id="441b3-103">Host element</span></span>

<span data-ttu-id="441b3-104">Especifica um tipo de aplicativo individual do Office em que o suplemento deve ser ativado.</span><span class="sxs-lookup"><span data-stu-id="441b3-104">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="441b3-105">A sintaxe do elemento **Host** varia de acordo com a definição do elemento, se dentro do [manifesto básico](#basic-manifest) ou dentro do nó [VersionOverrides](#versionoverrides-node).</span><span class="sxs-lookup"><span data-stu-id="441b3-105">The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="441b3-106">No entanto, a funcionalidade é a mesma.</span><span class="sxs-lookup"><span data-stu-id="441b3-106">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="441b3-107">Manifesto básico</span><span class="sxs-lookup"><span data-stu-id="441b3-107">Basic manifest</span></span>

<span data-ttu-id="441b3-108">Quando definido no manifesto básico (em [OfficeApp](officeapp.md)), o tipo de host é determinado pelo atributo `Name`.</span><span class="sxs-lookup"><span data-stu-id="441b3-108">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="441b3-109">Atributos</span><span class="sxs-lookup"><span data-stu-id="441b3-109">Attributes</span></span>

| <span data-ttu-id="441b3-110">Atributo</span><span class="sxs-lookup"><span data-stu-id="441b3-110">Attribute</span></span>     | <span data-ttu-id="441b3-111">Tipo</span><span class="sxs-lookup"><span data-stu-id="441b3-111">Type</span></span>   | <span data-ttu-id="441b3-112">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="441b3-112">Required</span></span> | <span data-ttu-id="441b3-113">Descrição</span><span class="sxs-lookup"><span data-stu-id="441b3-113">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="441b3-114">Nome</span><span class="sxs-lookup"><span data-stu-id="441b3-114">Name</span></span>](#name) | <span data-ttu-id="441b3-115">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="441b3-115">string</span></span> | <span data-ttu-id="441b3-116">obrigatório</span><span class="sxs-lookup"><span data-stu-id="441b3-116">required</span></span> | <span data-ttu-id="441b3-117">O nome do tipo de aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="441b3-117">The name of the type of Office host application.</span></span> |

### <a name="name"></a><span data-ttu-id="441b3-118">Name</span><span class="sxs-lookup"><span data-stu-id="441b3-118">Name</span></span>

<span data-ttu-id="441b3-119">Especifica o tipo de Host destinado por esse suplemento.</span><span class="sxs-lookup"><span data-stu-id="441b3-119">Specifies the Host type targeted by this add-in.</span></span> <span data-ttu-id="441b3-120">O valor deve ser um dos seguintes.</span><span class="sxs-lookup"><span data-stu-id="441b3-120">The value must be one of the following.</span></span>

- <span data-ttu-id="441b3-121">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="441b3-121">`Document` (Word)</span></span>
- <span data-ttu-id="441b3-122">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="441b3-122">`Database` (Access)</span></span>
- <span data-ttu-id="441b3-123">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="441b3-123">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="441b3-124">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="441b3-124">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="441b3-125">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="441b3-125">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="441b3-126">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="441b3-126">`Project` (Project)</span></span>
- <span data-ttu-id="441b3-127">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="441b3-127">`Workbook` (Excel)</span></span>

> [!IMPORTANT]
> <span data-ttu-id="441b3-128">Não recomendamos mais criar e usar aplicativos Web do Access e bancos de dados no SharePoint.</span><span class="sxs-lookup"><span data-stu-id="441b3-128">We no longer recommend that you create and use Access web apps and databases in SharePoint.</span></span> <span data-ttu-id="441b3-129">Como alternativa, use o [Microsoft PowerApps](https://powerapps.microsoft.com/) para criar soluções de negócios sem código para dispositivos móveis e Web.</span><span class="sxs-lookup"><span data-stu-id="441b3-129">As an alternative, we recommend that you use [Microsoft PowerApps](https://powerapps.microsoft.com/) to build no-code business solutions for web and mobile devices.</span></span>

### <a name="example"></a><span data-ttu-id="441b3-130">Exemplo</span><span class="sxs-lookup"><span data-stu-id="441b3-130">Example</span></span>

```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="441b3-131">Nó VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="441b3-131">VersionOverrides node</span></span>

<span data-ttu-id="441b3-132">Quando definido em [VersionOverrides](versionoverrides.md), o tipo de host é determinado pelo atributo `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="441b3-132">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="441b3-133">Atributos</span><span class="sxs-lookup"><span data-stu-id="441b3-133">Attributes</span></span>

|  <span data-ttu-id="441b3-134">Atributo</span><span class="sxs-lookup"><span data-stu-id="441b3-134">Attribute</span></span>  |  <span data-ttu-id="441b3-135">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="441b3-135">Required</span></span>  |  <span data-ttu-id="441b3-136">Descrição</span><span class="sxs-lookup"><span data-stu-id="441b3-136">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="441b3-137">xsi:type</span><span class="sxs-lookup"><span data-stu-id="441b3-137">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="441b3-138">Sim</span><span class="sxs-lookup"><span data-stu-id="441b3-138">Yes</span></span>  | <span data-ttu-id="441b3-139">Descreve o host do Office a que essas configurações se aplicam.</span><span class="sxs-lookup"><span data-stu-id="441b3-139">Describes the Office host where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="441b3-140">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="441b3-140">Child elements</span></span>

|  <span data-ttu-id="441b3-141">Elemento</span><span class="sxs-lookup"><span data-stu-id="441b3-141">Element</span></span> |  <span data-ttu-id="441b3-142">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="441b3-142">Required</span></span>  |  <span data-ttu-id="441b3-143">Descrição</span><span class="sxs-lookup"><span data-stu-id="441b3-143">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="441b3-144">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="441b3-144">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="441b3-145">Sim</span><span class="sxs-lookup"><span data-stu-id="441b3-145">Yes</span></span>   |  <span data-ttu-id="441b3-146">Define as configurações do fator forma da área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="441b3-146">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="441b3-147">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="441b3-147">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="441b3-148">Não</span><span class="sxs-lookup"><span data-stu-id="441b3-148">No</span></span>   |  <span data-ttu-id="441b3-149">Define as configurações do fator forma móvel.</span><span class="sxs-lookup"><span data-stu-id="441b3-149">Defines the settings for the mobile form factor.</span></span> <span data-ttu-id="441b3-150">**Observação:** Esse elemento só é suportado no Outlook no iOS e no Android.</span><span class="sxs-lookup"><span data-stu-id="441b3-150">**Note:** This element is only supported in Outlook on iOS and Android.</span></span> |
|  [<span data-ttu-id="441b3-151">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="441b3-151">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="441b3-152">Não</span><span class="sxs-lookup"><span data-stu-id="441b3-152">No</span></span>   |  <span data-ttu-id="441b3-153">Define as configurações de todos os fatores forma.</span><span class="sxs-lookup"><span data-stu-id="441b3-153">Defines the settings for all form factors.</span></span> <span data-ttu-id="441b3-154">Usado somente pelas funções personalizadas no Excel.</span><span class="sxs-lookup"><span data-stu-id="441b3-154">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="441b3-155">xsi:type</span><span class="sxs-lookup"><span data-stu-id="441b3-155">xsi:type</span></span>

<span data-ttu-id="441b3-156">Controla a qual host do Office (Word, Excel, PowerPoint, Outlook, OneNote) as configurações contidas se aplicam.</span><span class="sxs-lookup"><span data-stu-id="441b3-156">Controls which Office host (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="441b3-157">O valor deve ser uma das seguintes opções:</span><span class="sxs-lookup"><span data-stu-id="441b3-157">The value must be one of the following:</span></span>

- <span data-ttu-id="441b3-158">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="441b3-158">`Document` (Word)</span></span>
- <span data-ttu-id="441b3-159">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="441b3-159">`MailHost` (Outlook)</span></span>
- <span data-ttu-id="441b3-160">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="441b3-160">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="441b3-161">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="441b3-161">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="441b3-162">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="441b3-162">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="441b3-163">Exemplo de host</span><span class="sxs-lookup"><span data-stu-id="441b3-163">Host example</span></span>

```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
