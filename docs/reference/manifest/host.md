---
title: Elemento Host no arquivo de manifesto
description: ''
ms.date: 11/05/2019
localization_priority: Normal
ms.openlocfilehash: 824cc6ae51eb9db713a0a9a768e3ec48e3271e95
ms.sourcegitcommit: 08c0b9ff319c391922fa43d3c2e9783cf6b53b1b
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/08/2019
ms.locfileid: "38066274"
---
# <a name="host-element"></a><span data-ttu-id="cc322-102">Elemento Host</span><span class="sxs-lookup"><span data-stu-id="cc322-102">Host element</span></span>

<span data-ttu-id="cc322-103">Especifica um tipo de aplicativo individual do Office em que o suplemento deve ser ativado.</span><span class="sxs-lookup"><span data-stu-id="cc322-103">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="cc322-104">A sintaxe do elemento **Host** varia de acordo com a definição do elemento, se dentro do [manifesto básico](#basic-manifest) ou dentro do nó [VersionOverrides](#versionoverrides-node).</span><span class="sxs-lookup"><span data-stu-id="cc322-104">The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="cc322-105">No entanto, a funcionalidade é a mesma.</span><span class="sxs-lookup"><span data-stu-id="cc322-105">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="cc322-106">Manifesto básico</span><span class="sxs-lookup"><span data-stu-id="cc322-106">Basic manifest</span></span>

<span data-ttu-id="cc322-107">Quando definido no manifesto básico (em [OfficeApp](officeapp.md)), o tipo de host é determinado pelo atributo `Name`.</span><span class="sxs-lookup"><span data-stu-id="cc322-107">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="cc322-108">Atributos</span><span class="sxs-lookup"><span data-stu-id="cc322-108">Attributes</span></span>

| <span data-ttu-id="cc322-109">Atributo</span><span class="sxs-lookup"><span data-stu-id="cc322-109">Attribute</span></span>     | <span data-ttu-id="cc322-110">Tipo</span><span class="sxs-lookup"><span data-stu-id="cc322-110">Type</span></span>   | <span data-ttu-id="cc322-111">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="cc322-111">Required</span></span> | <span data-ttu-id="cc322-112">Descrição</span><span class="sxs-lookup"><span data-stu-id="cc322-112">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="cc322-113">Nome</span><span class="sxs-lookup"><span data-stu-id="cc322-113">Name</span></span>](#name) | <span data-ttu-id="cc322-114">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="cc322-114">string</span></span> | <span data-ttu-id="cc322-115">obrigatório</span><span class="sxs-lookup"><span data-stu-id="cc322-115">required</span></span> | <span data-ttu-id="cc322-116">O nome do tipo de aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="cc322-116">The name of the type of Office host application.</span></span> |

### <a name="name"></a><span data-ttu-id="cc322-117">Name</span><span class="sxs-lookup"><span data-stu-id="cc322-117">Name</span></span>

<span data-ttu-id="cc322-118">Especifica o tipo de Host destinado por esse suplemento.</span><span class="sxs-lookup"><span data-stu-id="cc322-118">Specifies the Host type targeted by this add-in.</span></span> <span data-ttu-id="cc322-119">O valor deve ser um dos seguintes.</span><span class="sxs-lookup"><span data-stu-id="cc322-119">The value must be one of the following.</span></span>

- <span data-ttu-id="cc322-120">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="cc322-120">`Document` (Word)</span></span>
- <span data-ttu-id="cc322-121">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="cc322-121">`Database` (Access)</span></span>
- <span data-ttu-id="cc322-122">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="cc322-122">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="cc322-123">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="cc322-123">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="cc322-124">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="cc322-124">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="cc322-125">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="cc322-125">`Project` (Project)</span></span>
- <span data-ttu-id="cc322-126">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="cc322-126">`Workbook` (Excel)</span></span>

> [!IMPORTANT]
> <span data-ttu-id="cc322-127">Não recomendamos mais criar e usar aplicativos Web do Access e bancos de dados no SharePoint.</span><span class="sxs-lookup"><span data-stu-id="cc322-127">We no longer recommend that you create and use Access web apps and databases in SharePoint.</span></span> <span data-ttu-id="cc322-128">Como alternativa, use o [Microsoft PowerApps](https://powerapps.microsoft.com/) para criar soluções de negócios sem código para dispositivos móveis e Web.</span><span class="sxs-lookup"><span data-stu-id="cc322-128">As an alternative, we recommend that you use [Microsoft PowerApps](https://powerapps.microsoft.com/) to build no-code business solutions for web and mobile devices.</span></span>

### <a name="example"></a><span data-ttu-id="cc322-129">Exemplo</span><span class="sxs-lookup"><span data-stu-id="cc322-129">Example</span></span>

```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="cc322-130">Nó VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="cc322-130">VersionOverrides node</span></span>

<span data-ttu-id="cc322-131">Quando definido em [VersionOverrides](versionoverrides.md), o tipo de host é determinado pelo atributo `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="cc322-131">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="cc322-132">Atributos</span><span class="sxs-lookup"><span data-stu-id="cc322-132">Attributes</span></span>

|  <span data-ttu-id="cc322-133">Atributo</span><span class="sxs-lookup"><span data-stu-id="cc322-133">Attribute</span></span>  |  <span data-ttu-id="cc322-134">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="cc322-134">Required</span></span>  |  <span data-ttu-id="cc322-135">Descrição</span><span class="sxs-lookup"><span data-stu-id="cc322-135">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="cc322-136">xsi:type</span><span class="sxs-lookup"><span data-stu-id="cc322-136">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="cc322-137">Sim</span><span class="sxs-lookup"><span data-stu-id="cc322-137">Yes</span></span>  | <span data-ttu-id="cc322-138">Descreve o host do Office a que essas configurações se aplicam.</span><span class="sxs-lookup"><span data-stu-id="cc322-138">Describes the Office host where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="cc322-139">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="cc322-139">Child elements</span></span>

|  <span data-ttu-id="cc322-140">Elemento</span><span class="sxs-lookup"><span data-stu-id="cc322-140">Element</span></span> |  <span data-ttu-id="cc322-141">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="cc322-141">Required</span></span>  |  <span data-ttu-id="cc322-142">Descrição</span><span class="sxs-lookup"><span data-stu-id="cc322-142">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="cc322-143">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="cc322-143">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="cc322-144">Sim</span><span class="sxs-lookup"><span data-stu-id="cc322-144">Yes</span></span>   |  <span data-ttu-id="cc322-145">Define as configurações do fator forma da área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="cc322-145">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="cc322-146">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="cc322-146">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="cc322-147">Não</span><span class="sxs-lookup"><span data-stu-id="cc322-147">No</span></span>   |  <span data-ttu-id="cc322-148">Define as configurações do fator forma móvel.</span><span class="sxs-lookup"><span data-stu-id="cc322-148">Defines the settings for the mobile form factor.</span></span> <span data-ttu-id="cc322-149">**Observação:** Esse elemento só é suportado no Outlook no iOS e no Android.</span><span class="sxs-lookup"><span data-stu-id="cc322-149">**Note:** This element is only supported in Outlook on iOS and Android.</span></span> |
|  [<span data-ttu-id="cc322-150">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="cc322-150">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="cc322-151">Não</span><span class="sxs-lookup"><span data-stu-id="cc322-151">No</span></span>   |  <span data-ttu-id="cc322-152">Define as configurações de todos os fatores forma.</span><span class="sxs-lookup"><span data-stu-id="cc322-152">Defines the settings for all form factors.</span></span> <span data-ttu-id="cc322-153">Usado somente pelas funções personalizadas no Excel.</span><span class="sxs-lookup"><span data-stu-id="cc322-153">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="cc322-154">xsi:type</span><span class="sxs-lookup"><span data-stu-id="cc322-154">xsi:type</span></span>

<span data-ttu-id="cc322-155">Controla a qual host do Office (Word, Excel, PowerPoint, Outlook, OneNote) as configurações contidas se aplicam.</span><span class="sxs-lookup"><span data-stu-id="cc322-155">Controls which Office host (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="cc322-156">O valor deve ser uma das seguintes opções:</span><span class="sxs-lookup"><span data-stu-id="cc322-156">The value must be one of the following:</span></span>

- <span data-ttu-id="cc322-157">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="cc322-157">`Document` (Word)</span></span>
- <span data-ttu-id="cc322-158">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="cc322-158">`MailHost` (Outlook)</span></span>
- <span data-ttu-id="cc322-159">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="cc322-159">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="cc322-160">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="cc322-160">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="cc322-161">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="cc322-161">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="cc322-162">Exemplo de host</span><span class="sxs-lookup"><span data-stu-id="cc322-162">Host example</span></span>

```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
