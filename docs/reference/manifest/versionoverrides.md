---
title: Elemento VersionOverrides no arquivo de manifesto
description: ''
ms.date: 02/04/2020
localization_priority: Normal
ms.openlocfilehash: 0afa3183e34a736a878217c079b7b8d0259be5b1
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324796"
---
# <a name="versionoverrides-element"></a><span data-ttu-id="12819-102">Elemento VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="12819-102">VersionOverrides element</span></span>

<span data-ttu-id="12819-p101">O elemento raiz que contém informações para os comandos de suplemento implementados pelo suplemento. **VersionOverrides** é um elemento filho do elemento [OfficeApp](./officeapp.md) no manifesto. Ele recebe suporte no esquema de manifesto v1.1 e posterior, mas é definido no esquema VersionOverrides v1.0 ou v1.1.</span><span class="sxs-lookup"><span data-stu-id="12819-p101">The root element that contains information for the add-in commands implemented by the add-in. **VersionOverrides** is a child element of the [OfficeApp](./officeapp.md) element in the manifest. This element is supported in manifest schema v1.1 and later but is defined in the VersionOverrides v1.0 or v1.1 schema.</span></span>

## <a name="attributes"></a><span data-ttu-id="12819-106">Atributos</span><span class="sxs-lookup"><span data-stu-id="12819-106">Attributes</span></span>

|  <span data-ttu-id="12819-107">Atributo</span><span class="sxs-lookup"><span data-stu-id="12819-107">Attribute</span></span>  |  <span data-ttu-id="12819-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="12819-108">Required</span></span>  |  <span data-ttu-id="12819-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="12819-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="12819-110">**xmlns**</span><span class="sxs-lookup"><span data-stu-id="12819-110">**xmlns**</span></span>       |  <span data-ttu-id="12819-111">Sim</span><span class="sxs-lookup"><span data-stu-id="12819-111">Yes</span></span>  |  <span data-ttu-id="12819-112">O namespace do esquema VersionOverrides.</span><span class="sxs-lookup"><span data-stu-id="12819-112">The VersionOverrides schema namespace.</span></span> <span data-ttu-id="12819-113">Os valores permitidos variam de acordo com o `<VersionOverrides>` valor **xsi: Type** do elemento e o valor **xsi: Type** do elemento pai `<OfficeApp>` .</span><span class="sxs-lookup"><span data-stu-id="12819-113">The allowed values vary depending on  this `<VersionOverrides>` element's **xsi:type** value and the **xsi:type** value of the parent `<OfficeApp>` element.</span></span> <span data-ttu-id="12819-114">Consulte [namespace valores](#namespace-values) a seguir.</span><span class="sxs-lookup"><span data-stu-id="12819-114">See [Namespace values](#namespace-values) below.</span></span>|
|  <span data-ttu-id="12819-115">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="12819-115">**xsi:type**</span></span>  |  <span data-ttu-id="12819-116">Sim</span><span class="sxs-lookup"><span data-stu-id="12819-116">Yes</span></span>  | <span data-ttu-id="12819-p103">A versão do esquema. Nesse momento, os únicos valores válidos são `VersionOverridesV1_0` e `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="12819-p103">The schema version. At this time, the only valid values are `VersionOverridesV1_0` and `VersionOverridesV1_1`.</span></span> |

### <a name="namespace-values"></a><span data-ttu-id="12819-119">Valores do namespace</span><span class="sxs-lookup"><span data-stu-id="12819-119">Namespace values</span></span>

<span data-ttu-id="12819-120">A seguir, a lista o valor necessário do valor **xmlns** , dependendo do valor **xsi: Type** do elemento pai `<OfficeApp>` .</span><span class="sxs-lookup"><span data-stu-id="12819-120">The following lists the required value of the **xmlns** value depending on the **xsi:type** value of the parent `<OfficeApp>` element.</span></span>

- <span data-ttu-id="12819-121">**TaskPaneApp** oferece suporte somente à versão 1,0 do VersionOverrides e o **xmlns** deve `http://schemas.microsoft.com/office/taskpaneappversionoverrides`ser.</span><span class="sxs-lookup"><span data-stu-id="12819-121">**TaskPaneApp** supports only version 1.0 of VersionOverrides, and the **xmlns** should be `http://schemas.microsoft.com/office/taskpaneappversionoverrides`.</span></span>
- <span data-ttu-id="12819-122">**ContentApp** oferece suporte somente à versão 1,0 do VersionOverrides e o **xmlns** deve `http://schemas.microsoft.com/office/contentappversionoverrides`ser.</span><span class="sxs-lookup"><span data-stu-id="12819-122">**ContentApp** supports only version 1.0 of VersionOverrides, and the **xmlns** should be `http://schemas.microsoft.com/office/contentappversionoverrides`.</span></span>
- <span data-ttu-id="12819-123">O **MailApp** suporta as versões 1,0 e 1,1 do VersionOverrides, portanto, o valor de **xmlns** varia de `<VersionOverrides>` acordo com o valor **xsi: Type** do elemento:</span><span class="sxs-lookup"><span data-stu-id="12819-123">**MailApp** supports versions 1.0 and 1.1 of VersionOverrides, so the value of **xmlns** varies depending on this `<VersionOverrides>` element's **xsi:type** value:</span></span>
    - <span data-ttu-id="12819-124">Quando **xsi: Type** é `VersionOverridesV1_0`, o **xmlns** deve `http://schemas.microsoft.com/office/mailappversionoverrides`ser.</span><span class="sxs-lookup"><span data-stu-id="12819-124">When **xsi:type** is `VersionOverridesV1_0`, **xmlns** must be `http://schemas.microsoft.com/office/mailappversionoverrides`.</span></span>
    - <span data-ttu-id="12819-125">Quando **xsi: Type** é `VersionOverridesV1_1`, o **xmlns** deve `http://schemas.microsoft.com/office/mailappversionoverrides/1.1`ser.</span><span class="sxs-lookup"><span data-stu-id="12819-125">When **xsi:type** is `VersionOverridesV1_1`, **xmlns** must be `http://schemas.microsoft.com/office/mailappversionoverrides/1.1`.</span></span>

> [!NOTE]
> <span data-ttu-id="12819-126">Atualmente, somente o Outlook 2016 ou posterior oferece suporte ao esquema do VersionOverrides `VersionOverridesV1_1` v 1.1 e ao tipo.</span><span class="sxs-lookup"><span data-stu-id="12819-126">Currently only Outlook 2016 or later supports the VersionOverrides v1.1 schema and the `VersionOverridesV1_1` type.</span></span>

## <a name="child-elements"></a><span data-ttu-id="12819-127">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="12819-127">Child elements</span></span>

|  <span data-ttu-id="12819-128">Elemento</span><span class="sxs-lookup"><span data-stu-id="12819-128">Element</span></span> |  <span data-ttu-id="12819-129">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="12819-129">Required</span></span>  |  <span data-ttu-id="12819-130">Descrição</span><span class="sxs-lookup"><span data-stu-id="12819-130">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="12819-131">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="12819-131">**Description**</span></span>    |  <span data-ttu-id="12819-132">Não</span><span class="sxs-lookup"><span data-stu-id="12819-132">No</span></span>   |  <span data-ttu-id="12819-p104">Descreve o suplemento. Isso substitui o elemento `Description` em qualquer parte pai do manifesto. O texto da descrição está contido em um elemento filho do elemento **LongString**, contido no elemento [Resources](./resources.md). O atributo `resid` do elemento **Description** está definido como o valor do atributo `id` do elemento `String` que contém o texto.</span><span class="sxs-lookup"><span data-stu-id="12819-p104">Describes the add-in. This overrides the `Description` element in any parent portion of the manifest. The text of the description is contained in a child element of the **LongString** element contained in the [Resources](./resources.md) element. The `resid` attribute of the **Description** element is set to the value of the `id` attribute of the `String` element that contains the text.</span></span>|
|  <span data-ttu-id="12819-137">**Requisitos**</span><span class="sxs-lookup"><span data-stu-id="12819-137">**Requirements**</span></span>  |  <span data-ttu-id="12819-138">Não</span><span class="sxs-lookup"><span data-stu-id="12819-138">No</span></span>   |  <span data-ttu-id="12819-p105">Especifica o conjunto de requisitos mínimos e a versão do Office.js exigida pelo suplemento. Isso substitui o elemento `Requirements` na parte pai do manifesto.</span><span class="sxs-lookup"><span data-stu-id="12819-p105">Specifies the minimum requirement set and version of Office.js that the add-in requires. This overrides the  `Requirements` element in the parent portion of the manifest.</span></span>|
|  [<span data-ttu-id="12819-141">Hosts</span><span class="sxs-lookup"><span data-stu-id="12819-141">Hosts</span></span>](./hosts.md)                |  <span data-ttu-id="12819-142">Sim</span><span class="sxs-lookup"><span data-stu-id="12819-142">Yes</span></span>  |  <span data-ttu-id="12819-p106">Especifica um conjunto de hosts do Office. O elemento filho Hosts substitui o elemento Hosts na parte pai do manifesto.</span><span class="sxs-lookup"><span data-stu-id="12819-p106">Specifies a collection of Office hosts. The child  Hosts element overrides the Hosts element in the parent portion of the manifest.</span></span>  |
|  [<span data-ttu-id="12819-145">Resources</span><span class="sxs-lookup"><span data-stu-id="12819-145">Resources</span></span>](./resources.md)    |  <span data-ttu-id="12819-146">Sim</span><span class="sxs-lookup"><span data-stu-id="12819-146">Yes</span></span>  | <span data-ttu-id="12819-147">Define um conjunto de recursos (cadeias de caracteres, URLs e imagens) consultado por outros elementos do manifesto.</span><span class="sxs-lookup"><span data-stu-id="12819-147">Defines a collection of resources (strings, URLs, and images) that other manifest elements reference.</span></span>|
|  [<span data-ttu-id="12819-148">EquivalentAddins</span><span class="sxs-lookup"><span data-stu-id="12819-148">EquivalentAddins</span></span>](./equivalentaddins.md)    |  <span data-ttu-id="12819-149">Não</span><span class="sxs-lookup"><span data-stu-id="12819-149">No</span></span>  | <span data-ttu-id="12819-150">Especifica os suplementos nativos (COM/XLL) equivalentes ao suplemento Web.</span><span class="sxs-lookup"><span data-stu-id="12819-150">Specifies the native (COM/XLL) add-ins that are equivalent to the web add-in.</span></span> <span data-ttu-id="12819-151">O suplemento Web não será ativado se um suplemento nativo equivalente estiver instalado.</span><span class="sxs-lookup"><span data-stu-id="12819-151">The web add-in is not activated if an equivalent native add-in is installed.</span></span>|
|  <span data-ttu-id="12819-152">**VersionOverrides**</span><span class="sxs-lookup"><span data-stu-id="12819-152">**VersionOverrides**</span></span>    |  <span data-ttu-id="12819-153">Não</span><span class="sxs-lookup"><span data-stu-id="12819-153">No</span></span>  | <span data-ttu-id="12819-p108">Define comandos de suplemento em uma versão mais recente do esquema. Para saber mais, confira o tópico [Implementar várias versões](#implementing-multiple-versions).</span><span class="sxs-lookup"><span data-stu-id="12819-p108">Defines add-in commands under a newer schema version. See [Implementing multiple versions](#implementing-multiple-versions) for details.</span></span> |
|  [<span data-ttu-id="12819-156">WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="12819-156">WebApplicationInfo</span></span>](./webapplicationinfo.md)    |  <span data-ttu-id="12819-157">Não</span><span class="sxs-lookup"><span data-stu-id="12819-157">No</span></span>  | <span data-ttu-id="12819-158">Especifica detalhes sobre o registro do suplemento com emissores de token seguros, como o Azure Active Directory V 2.0.</span><span class="sxs-lookup"><span data-stu-id="12819-158">Specifies details about the add-in's registration with secure token issuers, such as Azure Active Directory V2.0.</span></span> |

### <a name="versionoverrides-example"></a><span data-ttu-id="12819-159">Exemplo de VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="12819-159">VersionOverrides example</span></span>

<span data-ttu-id="12819-160">Veja a seguir um exemplo de um elemento `<VersionOverrides>` típico, incluindo alguns elementos filhos que não são necessários, mas que são normalmente usados.</span><span class="sxs-lookup"><span data-stu-id="12819-160">The following is an example of a typical `<VersionOverrides>` element, including some child elements that are not required but are typically used.</span></span>

```xml
<OfficeApp ... xsi:type="MailApp">
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources>
      <!-- add information on resources -->
    </Resources>
  </VersionOverrides>
...
</OfficeApp>
```

## <a name="implementing-multiple-versions"></a><span data-ttu-id="12819-161">Implementar várias versões</span><span class="sxs-lookup"><span data-stu-id="12819-161">Implementing multiple versions</span></span>

<span data-ttu-id="12819-p109">Um manifesto pode implementar várias versões do elemento `VersionOverrides` que é compatível com várias versões do esquema VersionOverrides. Isso pode ser feito para fornecer suporte opcional a novos recursos em um esquema mais recente, sem deixar de fornecer suporte a clientes antigos que não têm suporte para os novos recursos.</span><span class="sxs-lookup"><span data-stu-id="12819-p109">A manifest can implement multiple versions of the `VersionOverrides` element which support different versions of the VersionOverrides schema. This can be done to optionally support new features in a newer schema while still supporting older clients that do not support the new features.</span></span>

<span data-ttu-id="12819-164">Para implementar várias versões, o elemento `VersionOverrides` da versão mais recente deve ser um filho do elemento `VersionOverrides` da versão anterior.</span><span class="sxs-lookup"><span data-stu-id="12819-164">In order to implement multiple versions, the `VersionOverrides` element for the newer version must be a child of the `VersionOverrides` element for the older version.</span></span> <span data-ttu-id="12819-165">O elemento filho `VersionOverrides` não herda os valores do elemento pai.</span><span class="sxs-lookup"><span data-stu-id="12819-165">The child `VersionOverrides` element doesn't inherit any values from the parent.</span></span>

<span data-ttu-id="12819-166">Para implementar o esquema do VersionOverrides v1.0 e do v1.1, o manifesto seria semelhante ao exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="12819-166">To implement both the VersionOverrides v1.0 and v1.1 schema, the manifest would look similar to the following example:</span></span>

```xml
<OfficeApp ... xsi:type="MailApp">
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources>
      <!-- add information on resources -->
    </Resources>

    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
      <Description resid="residDescription" />
      <Requirements>
        <!-- add information on requirements -->
      </Requirements>
      <Hosts>
        <Host xsi:type="MailHost">
          <!-- add information on form factors -->
        </Host>
      </Hosts>
      <Resources>
        <!-- add information on resources -->
      </Resources>
    </VersionOverrides>  
  </VersionOverrides>
...
</OfficeApp>
```
