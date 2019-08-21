---
title: Elemento VersionOverrides no arquivo de manifesto
description: ''
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: ce65cdced1b3cf885cee09732c2cda0081a53cfc
ms.sourcegitcommit: da8e6148f4bd9884ab9702db3033273a383d15f0
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/20/2019
ms.locfileid: "36477877"
---
# <a name="versionoverrides-element"></a><span data-ttu-id="4ce7d-102">Elemento VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="4ce7d-102">VersionOverrides element</span></span>

<span data-ttu-id="4ce7d-p101">O elemento raiz que contém informações para os comandos de suplemento implementados pelo suplemento. **VersionOverrides** é um elemento filho do elemento [OfficeApp](./officeapp.md) no manifesto. Ele recebe suporte no esquema de manifesto v1.1 e posterior, mas é definido no esquema VersionOverrides v1.0 ou v1.1.</span><span class="sxs-lookup"><span data-stu-id="4ce7d-p101">The root element that contains information for the add-in commands implemented by the add-in. **VersionOverrides** is a child element of the [OfficeApp](./officeapp.md) element in the manifest. This element is supported in manifest schema v1.1 and later but is defined in the VersionOverrides v1.0 or v1.1 schema.</span></span>

## <a name="attributes"></a><span data-ttu-id="4ce7d-106">Atributos</span><span class="sxs-lookup"><span data-stu-id="4ce7d-106">Attributes</span></span>

|  <span data-ttu-id="4ce7d-107">Atributo</span><span class="sxs-lookup"><span data-stu-id="4ce7d-107">Attribute</span></span>  |  <span data-ttu-id="4ce7d-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="4ce7d-108">Required</span></span>  |  <span data-ttu-id="4ce7d-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="4ce7d-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="4ce7d-110">**xmlns**</span><span class="sxs-lookup"><span data-stu-id="4ce7d-110">**xmlns**</span></span>       |  <span data-ttu-id="4ce7d-111">Sim</span><span class="sxs-lookup"><span data-stu-id="4ce7d-111">Yes</span></span>  |  <span data-ttu-id="4ce7d-112">O local do esquema, que deve ser `http://schemas.microsoft.com/office/mailappversionoverrides` quando `xsi:type` for `VersionOverridesV1_0` e `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` quando `xsi:type` for `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="4ce7d-112">The schema location, which must be `http://schemas.microsoft.com/office/mailappversionoverrides` when `xsi:type` is `VersionOverridesV1_0`, and `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` when `xsi:type` is `VersionOverridesV1_1`.</span></span>|
|  <span data-ttu-id="4ce7d-113">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="4ce7d-113">**xsi:type**</span></span>  |  <span data-ttu-id="4ce7d-114">Sim</span><span class="sxs-lookup"><span data-stu-id="4ce7d-114">Yes</span></span>  | <span data-ttu-id="4ce7d-p102">A versão do esquema. Nesse momento, os únicos valores válidos são `VersionOverridesV1_0` e `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="4ce7d-p102">The schema version. At this time, the only valid values are `VersionOverridesV1_0` and `VersionOverridesV1_1`.</span></span> |

> [!NOTE]
> <span data-ttu-id="4ce7d-117">Atualmente, somente o Outlook 2016 ou posterior oferece suporte ao esquema do VersionOverrides `VersionOverridesV1_1` v 1.1 e ao tipo.</span><span class="sxs-lookup"><span data-stu-id="4ce7d-117">Currently only Outlook 2016 or later supports the VersionOverrides v1.1 schema and the `VersionOverridesV1_1` type.</span></span>

## <a name="child-elements"></a><span data-ttu-id="4ce7d-118">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="4ce7d-118">Child elements</span></span>

|  <span data-ttu-id="4ce7d-119">Elemento</span><span class="sxs-lookup"><span data-stu-id="4ce7d-119">Element</span></span> |  <span data-ttu-id="4ce7d-120">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="4ce7d-120">Required</span></span>  |  <span data-ttu-id="4ce7d-121">Descrição</span><span class="sxs-lookup"><span data-stu-id="4ce7d-121">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="4ce7d-122">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="4ce7d-122">**Description**</span></span>    |  <span data-ttu-id="4ce7d-123">Não</span><span class="sxs-lookup"><span data-stu-id="4ce7d-123">No</span></span>   |  <span data-ttu-id="4ce7d-p103">Descreve o suplemento. Isso substitui o elemento `Description` em qualquer parte pai do manifesto. O texto da descrição está contido em um elemento filho do elemento **LongString**, contido no elemento [Resources](./resources.md). O atributo `resid` do elemento **Description** está definido como o valor do atributo `id` do elemento `String` que contém o texto.</span><span class="sxs-lookup"><span data-stu-id="4ce7d-p103">Describes the add-in. This overrides the `Description` element in any parent portion of the manifest. The text of the description is contained in a child element of the **LongString** element contained in the [Resources](./resources.md) element. The `resid` attribute of the **Description** element is set to the value of the `id` attribute of the `String` element that contains the text.</span></span>|
| <span data-ttu-id="4ce7d-128">**EquivalentAddins**</span><span class="sxs-lookup"><span data-stu-id="4ce7d-128">**EquivalentAddins**</span></span> | <span data-ttu-id="4ce7d-129">Não</span><span class="sxs-lookup"><span data-stu-id="4ce7d-129">No</span></span> | <span data-ttu-id="4ce7d-130">Especifica a compatibilidade com versões anteriores com um suplemento COM equivalente, XLL ou ambos.</span><span class="sxs-lookup"><span data-stu-id="4ce7d-130">Specifies backwards compatibility with an equivalent COM add-in, XLL, or both.</span></span> |
|  <span data-ttu-id="4ce7d-131">**Requisitos**</span><span class="sxs-lookup"><span data-stu-id="4ce7d-131">**Requirements**</span></span>  |  <span data-ttu-id="4ce7d-132">Não</span><span class="sxs-lookup"><span data-stu-id="4ce7d-132">No</span></span>   |  <span data-ttu-id="4ce7d-p104">Especifica o conjunto de requisitos mínimos e a versão do Office.js exigida pelo suplemento. Isso substitui o elemento `Requirements` na parte pai do manifesto.</span><span class="sxs-lookup"><span data-stu-id="4ce7d-p104">Specifies the minimum requirement set and version of Office.js that the add-in requires. This overrides the  `Requirements` element in the parent portion of the manifest.</span></span>|
|  [<span data-ttu-id="4ce7d-135">Hosts</span><span class="sxs-lookup"><span data-stu-id="4ce7d-135">Hosts</span></span>](./hosts.md)                |  <span data-ttu-id="4ce7d-136">Sim</span><span class="sxs-lookup"><span data-stu-id="4ce7d-136">Yes</span></span>  |  <span data-ttu-id="4ce7d-p105">Especifica um conjunto de hosts do Office. O elemento filho Hosts substitui o elemento Hosts na parte pai do manifesto.</span><span class="sxs-lookup"><span data-stu-id="4ce7d-p105">Specifies a collection of Office hosts. The child  Hosts element overrides the Hosts element in the parent portion of the manifest.</span></span>  |
|  [<span data-ttu-id="4ce7d-139">Resources</span><span class="sxs-lookup"><span data-stu-id="4ce7d-139">Resources</span></span>](./resources.md)    |  <span data-ttu-id="4ce7d-140">Sim</span><span class="sxs-lookup"><span data-stu-id="4ce7d-140">Yes</span></span>  | <span data-ttu-id="4ce7d-141">Define um conjunto de recursos (cadeias de caracteres, URLs e imagens) consultado por outros elementos do manifesto.</span><span class="sxs-lookup"><span data-stu-id="4ce7d-141">Defines a collection of resources (strings, URLs, and images) that other manifest elements reference.</span></span>|
|  [<span data-ttu-id="4ce7d-142">EquivalentAddins</span><span class="sxs-lookup"><span data-stu-id="4ce7d-142">EquivalentAddins</span></span>](./equivalentaddins.md)    |  <span data-ttu-id="4ce7d-143">Não</span><span class="sxs-lookup"><span data-stu-id="4ce7d-143">No</span></span>  | <span data-ttu-id="4ce7d-144">Especifica os suplementos nativos (COM/XLL) equivalentes ao suplemento Web.</span><span class="sxs-lookup"><span data-stu-id="4ce7d-144">Specifies the native (COM/XLL) add-ins that are equivalent to the web add-in.</span></span> <span data-ttu-id="4ce7d-145">O suplemento Web não será ativado se um suplemento nativo equivalente estiver instalado.</span><span class="sxs-lookup"><span data-stu-id="4ce7d-145">The web add-in is not activated if an equivalent native add-in is installed.</span></span>|
|  <span data-ttu-id="4ce7d-146">**VersionOverrides**</span><span class="sxs-lookup"><span data-stu-id="4ce7d-146">**VersionOverrides**</span></span>    |  <span data-ttu-id="4ce7d-147">Não</span><span class="sxs-lookup"><span data-stu-id="4ce7d-147">No</span></span>  | <span data-ttu-id="4ce7d-p107">Define comandos de suplemento em uma versão mais recente do esquema. Para saber mais, confira o tópico [Implementar várias versões](#implementing-multiple-versions).</span><span class="sxs-lookup"><span data-stu-id="4ce7d-p107">Defines add-in commands under a newer schema version. See [Implementing multiple versions](#implementing-multiple-versions) for details.</span></span> |
|  [<span data-ttu-id="4ce7d-150">WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="4ce7d-150">WebApplicationInfo</span></span>](./webapplicationinfo.md)    |  <span data-ttu-id="4ce7d-151">Não</span><span class="sxs-lookup"><span data-stu-id="4ce7d-151">No</span></span>  | <span data-ttu-id="4ce7d-152">Especifica detalhes sobre o registro do suplemento com emissores de token seguros, como o Azure Active Directory V 2.0.</span><span class="sxs-lookup"><span data-stu-id="4ce7d-152">Specifies details about the add-in's registration with secure token issuers, such as Azure Active Directory V2.0.</span></span> |

### <a name="versionoverrides-example"></a><span data-ttu-id="4ce7d-153">Exemplo de VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="4ce7d-153">VersionOverrides example</span></span>

<span data-ttu-id="4ce7d-154">Veja a seguir um exemplo de um elemento `<VersionOverrides>` típico, incluindo alguns elementos filhos que não são necessários, mas que são normalmente usados.</span><span class="sxs-lookup"><span data-stu-id="4ce7d-154">The following is an example of a typical `<VersionOverrides>` element, including some child elements that are not required but are typically used.</span></span>

```xml
<OfficeApp>
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

## <a name="implementing-multiple-versions"></a><span data-ttu-id="4ce7d-155">Implementar várias versões</span><span class="sxs-lookup"><span data-stu-id="4ce7d-155">Implementing multiple versions</span></span>

<span data-ttu-id="4ce7d-p108">Um manifesto pode implementar várias versões do elemento `VersionOverrides` que é compatível com várias versões do esquema VersionOverrides. Isso pode ser feito para fornecer suporte opcional a novos recursos em um esquema mais recente, sem deixar de fornecer suporte a clientes antigos que não têm suporte para os novos recursos.</span><span class="sxs-lookup"><span data-stu-id="4ce7d-p108">A manifest can implement multiple versions of the `VersionOverrides` element which support different versions of the VersionOverrides schema. This can be done to optionally support new features in a newer schema while still supporting older clients that do not support the new features.</span></span>

<span data-ttu-id="4ce7d-158">Para implementar várias versões, o elemento `VersionOverrides` da versão mais recente deve ser um filho do elemento `VersionOverrides` da versão anterior.</span><span class="sxs-lookup"><span data-stu-id="4ce7d-158">In order to implement multiple versions, the `VersionOverrides` element for the newer version must be a child of the `VersionOverrides` element for the older version.</span></span> <span data-ttu-id="4ce7d-159">O elemento filho `VersionOverrides` não herda os valores do elemento pai.</span><span class="sxs-lookup"><span data-stu-id="4ce7d-159">The child `VersionOverrides` element doesn't inherit any values from the parent.</span></span>

<span data-ttu-id="4ce7d-160">Para implementar o esquema do VersionOverrides v1.0 e do v1.1, o manifesto seria semelhante ao exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="4ce7d-160">To implement both the VersionOverrides v1.0 and v1.1 schema, the manifest would look similar to the following example:</span></span>

```xml
<OfficeApp>
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
