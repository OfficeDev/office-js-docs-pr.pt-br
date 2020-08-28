---
title: Elemento Host no arquivo de manifesto
description: Especifica um tipo de aplicativo individual do Office em que o suplemento deve ser ativado.
ms.date: 11/05/2019
localization_priority: Normal
ms.openlocfilehash: 5b6c6e6b5471b4117c28cf92e11eb0a99b512a97
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292283"
---
# <a name="host-element"></a><span data-ttu-id="9cc92-103">Elemento Host</span><span class="sxs-lookup"><span data-stu-id="9cc92-103">Host element</span></span>

<span data-ttu-id="9cc92-104">Especifica um tipo de aplicativo individual do Office em que o suplemento deve ser ativado.</span><span class="sxs-lookup"><span data-stu-id="9cc92-104">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9cc92-105">A sintaxe do elemento **Host** varia de acordo com a definição do elemento, se dentro do [manifesto básico](#basic-manifest) ou dentro do nó [VersionOverrides](#versionoverrides-node).</span><span class="sxs-lookup"><span data-stu-id="9cc92-105">The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="9cc92-106">No entanto, a funcionalidade é a mesma.</span><span class="sxs-lookup"><span data-stu-id="9cc92-106">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="9cc92-107">Manifesto básico</span><span class="sxs-lookup"><span data-stu-id="9cc92-107">Basic manifest</span></span>

<span data-ttu-id="9cc92-108">Quando definido no manifesto básico (em [OfficeApp](officeapp.md)), o tipo de host é determinado pelo atributo `Name`.</span><span class="sxs-lookup"><span data-stu-id="9cc92-108">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="9cc92-109">Atributos</span><span class="sxs-lookup"><span data-stu-id="9cc92-109">Attributes</span></span>

| <span data-ttu-id="9cc92-110">Atributo</span><span class="sxs-lookup"><span data-stu-id="9cc92-110">Attribute</span></span>     | <span data-ttu-id="9cc92-111">Tipo</span><span class="sxs-lookup"><span data-stu-id="9cc92-111">Type</span></span>   | <span data-ttu-id="9cc92-112">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="9cc92-112">Required</span></span> | <span data-ttu-id="9cc92-113">Descrição</span><span class="sxs-lookup"><span data-stu-id="9cc92-113">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="9cc92-114">Nome</span><span class="sxs-lookup"><span data-stu-id="9cc92-114">Name</span></span>](#name) | <span data-ttu-id="9cc92-115">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9cc92-115">string</span></span> | <span data-ttu-id="9cc92-116">obrigatório</span><span class="sxs-lookup"><span data-stu-id="9cc92-116">required</span></span> | <span data-ttu-id="9cc92-117">O nome do tipo de aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="9cc92-117">The name of the type of Office client application.</span></span> |

### <a name="name"></a><span data-ttu-id="9cc92-118">Name</span><span class="sxs-lookup"><span data-stu-id="9cc92-118">Name</span></span>

<span data-ttu-id="9cc92-119">Especifica o tipo de Host destinado por esse suplemento.</span><span class="sxs-lookup"><span data-stu-id="9cc92-119">Specifies the Host type targeted by this add-in.</span></span> <span data-ttu-id="9cc92-120">O valor deve ser um dos seguintes.</span><span class="sxs-lookup"><span data-stu-id="9cc92-120">The value must be one of the following.</span></span>

- <span data-ttu-id="9cc92-121">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="9cc92-121">`Document` (Word)</span></span>
- <span data-ttu-id="9cc92-122">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="9cc92-122">`Database` (Access)</span></span>
- <span data-ttu-id="9cc92-123">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="9cc92-123">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="9cc92-124">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="9cc92-124">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="9cc92-125">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="9cc92-125">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="9cc92-126">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="9cc92-126">`Project` (Project)</span></span>
- <span data-ttu-id="9cc92-127">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="9cc92-127">`Workbook` (Excel)</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9cc92-128">Não recomendamos mais criar e usar aplicativos Web do Access e bancos de dados no SharePoint.</span><span class="sxs-lookup"><span data-stu-id="9cc92-128">We no longer recommend that you create and use Access web apps and databases in SharePoint.</span></span> <span data-ttu-id="9cc92-129">Como alternativa, use o [Microsoft PowerApps](https://powerapps.microsoft.com/) para criar soluções de negócios sem código para dispositivos móveis e Web.</span><span class="sxs-lookup"><span data-stu-id="9cc92-129">As an alternative, we recommend that you use [Microsoft PowerApps](https://powerapps.microsoft.com/) to build no-code business solutions for web and mobile devices.</span></span>

### <a name="example"></a><span data-ttu-id="9cc92-130">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9cc92-130">Example</span></span>

```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="9cc92-131">Nó VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="9cc92-131">VersionOverrides node</span></span>

<span data-ttu-id="9cc92-132">Quando definido em [VersionOverrides](versionoverrides.md), o tipo de host é determinado pelo atributo `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="9cc92-132">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="9cc92-133">Atributos</span><span class="sxs-lookup"><span data-stu-id="9cc92-133">Attributes</span></span>

|  <span data-ttu-id="9cc92-134">Atributo</span><span class="sxs-lookup"><span data-stu-id="9cc92-134">Attribute</span></span>  |  <span data-ttu-id="9cc92-135">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="9cc92-135">Required</span></span>  |  <span data-ttu-id="9cc92-136">Descrição</span><span class="sxs-lookup"><span data-stu-id="9cc92-136">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="9cc92-137">xsi:type</span><span class="sxs-lookup"><span data-stu-id="9cc92-137">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="9cc92-138">Sim</span><span class="sxs-lookup"><span data-stu-id="9cc92-138">Yes</span></span>  | <span data-ttu-id="9cc92-139">Descreve o aplicativo do Office onde essas configurações se aplicam.</span><span class="sxs-lookup"><span data-stu-id="9cc92-139">Describes the Office application where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="9cc92-140">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="9cc92-140">Child elements</span></span>

|  <span data-ttu-id="9cc92-141">Elemento</span><span class="sxs-lookup"><span data-stu-id="9cc92-141">Element</span></span> |  <span data-ttu-id="9cc92-142">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="9cc92-142">Required</span></span>  |  <span data-ttu-id="9cc92-143">Descrição</span><span class="sxs-lookup"><span data-stu-id="9cc92-143">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="9cc92-144">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="9cc92-144">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="9cc92-145">Sim</span><span class="sxs-lookup"><span data-stu-id="9cc92-145">Yes</span></span>   |  <span data-ttu-id="9cc92-146">Define as configurações do fator forma da área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="9cc92-146">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="9cc92-147">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="9cc92-147">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="9cc92-148">Não</span><span class="sxs-lookup"><span data-stu-id="9cc92-148">No</span></span>   |  <span data-ttu-id="9cc92-149">Define as configurações do fator forma móvel.</span><span class="sxs-lookup"><span data-stu-id="9cc92-149">Defines the settings for the mobile form factor.</span></span> <span data-ttu-id="9cc92-150">**Observação:** Esse elemento só é suportado no Outlook no iOS e no Android.</span><span class="sxs-lookup"><span data-stu-id="9cc92-150">**Note:** This element is only supported in Outlook on iOS and Android.</span></span> |
|  [<span data-ttu-id="9cc92-151">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="9cc92-151">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="9cc92-152">Não</span><span class="sxs-lookup"><span data-stu-id="9cc92-152">No</span></span>   |  <span data-ttu-id="9cc92-153">Define as configurações de todos os fatores forma.</span><span class="sxs-lookup"><span data-stu-id="9cc92-153">Defines the settings for all form factors.</span></span> <span data-ttu-id="9cc92-154">Usado somente pelas funções personalizadas no Excel.</span><span class="sxs-lookup"><span data-stu-id="9cc92-154">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="9cc92-155">xsi:type</span><span class="sxs-lookup"><span data-stu-id="9cc92-155">xsi:type</span></span>

<span data-ttu-id="9cc92-156">Controla qual aplicativo do Office (Word, Excel, PowerPoint, Outlook, OneNote) onde as configurações contidas se aplicam.</span><span class="sxs-lookup"><span data-stu-id="9cc92-156">Controls which Office application (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="9cc92-157">O valor deve ser uma das seguintes opções:</span><span class="sxs-lookup"><span data-stu-id="9cc92-157">The value must be one of the following:</span></span>

- <span data-ttu-id="9cc92-158">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="9cc92-158">`Document` (Word)</span></span>
- <span data-ttu-id="9cc92-159">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="9cc92-159">`MailHost` (Outlook)</span></span>
- <span data-ttu-id="9cc92-160">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="9cc92-160">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="9cc92-161">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="9cc92-161">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="9cc92-162">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="9cc92-162">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="9cc92-163">Exemplo de host</span><span class="sxs-lookup"><span data-stu-id="9cc92-163">Host example</span></span>

```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
