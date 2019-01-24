---
title: Elemento VersionOverrides no arquivo de manifesto
description: ''
ms.date: 01/15/2019
localization_priority: Normal
ms.openlocfilehash: 197a636169b7f00edd44019cee21686065845800
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/23/2019
ms.locfileid: "29387797"
---
# <a name="versionoverrides-element"></a><span data-ttu-id="a384f-102">Elemento VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="a384f-102">VersionOverrides element</span></span>

<span data-ttu-id="a384f-p101">O elemento raiz que contém informações para os comandos de suplemento implementados pelo suplemento. **VersionOverrides** é um elemento filho do elemento [OfficeApp](./officeapp.md) no manifesto. Ele recebe suporte no esquema de manifesto v1.1 e posterior, mas é definido no esquema VersionOverrides v1.0 ou v1.1.</span><span class="sxs-lookup"><span data-stu-id="a384f-p101">The root element that contains information for the add-in commands implemented by the add-in. **VersionOverrides** is a child element of the [OfficeApp](./officeapp.md) element in the manifest. This element is supported in manifest schema v1.1 and later but is defined in the VersionOverrides v1.0 or v1.1 schema.</span></span>

## <a name="attributes"></a><span data-ttu-id="a384f-106">Atributos</span><span class="sxs-lookup"><span data-stu-id="a384f-106">Attributes</span></span>

|  <span data-ttu-id="a384f-107">Atributo</span><span class="sxs-lookup"><span data-stu-id="a384f-107">Attribute</span></span>  |  <span data-ttu-id="a384f-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="a384f-108">Required</span></span>  |  <span data-ttu-id="a384f-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="a384f-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="a384f-110">**xmlns**</span><span class="sxs-lookup"><span data-stu-id="a384f-110">**xmlns**</span></span>       |  <span data-ttu-id="a384f-111">Sim</span><span class="sxs-lookup"><span data-stu-id="a384f-111">Yes</span></span>  |  <span data-ttu-id="a384f-112">O local do esquema, que deve ser `http://schemas.microsoft.com/office/mailappversionoverrides` quando `xsi:type` for `VersionOverridesV1_0` e `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` quando `xsi:type` for `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="a384f-112">The schema location, which must be `http://schemas.microsoft.com/office/mailappversionoverrides` when `xsi:type` is `VersionOverridesV1_0`, and `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` when `xsi:type` is `VersionOverridesV1_1`.</span></span>|
|  <span data-ttu-id="a384f-113">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="a384f-113">**xsi:type**</span></span>  |  <span data-ttu-id="a384f-114">Sim</span><span class="sxs-lookup"><span data-stu-id="a384f-114">Yes</span></span>  | <span data-ttu-id="a384f-p102">A versão do esquema. Nesse momento, os únicos valores válidos são `VersionOverridesV1_0` e `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="a384f-p102">The schema version. At this time, the only valid values are `VersionOverridesV1_0` and `VersionOverridesV1_1`.</span></span> |

> [!NOTE]
> <span data-ttu-id="a384f-117">Atualmente, apenas o Outlook 2016 é compatível com o esquema do VersionOverrides v1.1 e com o tipo `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="a384f-117">Currently only Outlook 2016 supports the VersionOverrides v1.1 schema and the `VersionOverridesV1_1` type.</span></span>

## <a name="child-elements"></a><span data-ttu-id="a384f-118">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="a384f-118">Child elements</span></span>

|  <span data-ttu-id="a384f-119">Elemento</span><span class="sxs-lookup"><span data-stu-id="a384f-119">Element</span></span> |  <span data-ttu-id="a384f-120">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="a384f-120">Required</span></span>  |  <span data-ttu-id="a384f-121">Descrição</span><span class="sxs-lookup"><span data-stu-id="a384f-121">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="a384f-122">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="a384f-122">**Description**</span></span>    |  <span data-ttu-id="a384f-123">Não</span><span class="sxs-lookup"><span data-stu-id="a384f-123">No</span></span>   |  <span data-ttu-id="a384f-p103">Descreve o suplemento. Isso substitui o elemento `Description` em qualquer parte pai do manifesto. O texto da descrição está contido em um elemento filho do elemento **LongString**, contido no elemento [Resources](./resources.md). O atributo `resid` do elemento **Description** está definido como o valor do atributo `id` do elemento `String` que contém o texto.</span><span class="sxs-lookup"><span data-stu-id="a384f-p103">Describes the add-in. This overrides the `Description` element in any parent portion of the manifest. The text of the description is contained in a child element of the **LongString** element contained in the [Resources](./resources.md) element. The `resid` attribute of the **Description** element is set to the value of the `id` attribute of the `String` element that contains the text.</span></span>|
|  <span data-ttu-id="a384f-128">**Requisitos**</span><span class="sxs-lookup"><span data-stu-id="a384f-128">**Requirements**</span></span>  |  <span data-ttu-id="a384f-129">Não</span><span class="sxs-lookup"><span data-stu-id="a384f-129">No</span></span>   |  <span data-ttu-id="a384f-p104">Especifica o conjunto de requisitos mínimos e a versão do Office.js exigida pelo suplemento. Isso substitui o elemento `Requirements` na parte pai do manifesto.</span><span class="sxs-lookup"><span data-stu-id="a384f-p104">Specifies the minimum requirement set and version of Office.js that the add-in requires. This overrides the  `Requirements` element in the parent portion of the manifest.</span></span>|
|  [<span data-ttu-id="a384f-132">Hosts</span><span class="sxs-lookup"><span data-stu-id="a384f-132">Hosts</span></span>](./hosts.md)                |  <span data-ttu-id="a384f-133">Sim</span><span class="sxs-lookup"><span data-stu-id="a384f-133">Yes</span></span>  |  <span data-ttu-id="a384f-p105">Especifica um conjunto de hosts do Office. O elemento filho Hosts substitui o elemento Hosts na parte pai do manifesto.</span><span class="sxs-lookup"><span data-stu-id="a384f-p105">Specifies a collection of Office hosts. The child  Hosts element overrides the Hosts element in the parent portion of the manifest.</span></span>  |
|  [<span data-ttu-id="a384f-136">Recursos</span><span class="sxs-lookup"><span data-stu-id="a384f-136">Resources</span></span>](./resources.md)    |  <span data-ttu-id="a384f-137">Sim</span><span class="sxs-lookup"><span data-stu-id="a384f-137">Yes</span></span>  | <span data-ttu-id="a384f-138">Define um conjunto de recursos (cadeias de caracteres, URLs e imagens) consultado por outros elementos do manifesto.</span><span class="sxs-lookup"><span data-stu-id="a384f-138">Defines a collection of resources (strings, URLs, and images) that other manifest elements reference.</span></span>|
|  <span data-ttu-id="a384f-139">**VersionOverrides**</span><span class="sxs-lookup"><span data-stu-id="a384f-139">**VersionOverrides**</span></span>    |  <span data-ttu-id="a384f-140">Não</span><span class="sxs-lookup"><span data-stu-id="a384f-140">No</span></span>  | <span data-ttu-id="a384f-p106">Define comandos de suplemento em uma versão mais recente do esquema. Para saber mais, confira o tópico [Implementar várias versões](#implementing-multiple-versions).</span><span class="sxs-lookup"><span data-stu-id="a384f-p106">Defines add-in commands under a newer schema version. See [Implementing multiple versions](#implementing-multiple-versions) for details.</span></span> |
|  <span data-ttu-id="a384f-143">**WebApplicationInfo**</span><span class="sxs-lookup"><span data-stu-id="a384f-143">**WebApplicationInfo**</span></span>    |  <span data-ttu-id="a384f-144">Não</span><span class="sxs-lookup"><span data-stu-id="a384f-144">No</span></span>  | <span data-ttu-id="a384f-145">Especifica detalhes sobre o aplicativo Web associado do suplemento.</span><span class="sxs-lookup"><span data-stu-id="a384f-145">Specifies details about the add-in's associated Web application.</span></span> |

### <a name="versionoverrides-example"></a><span data-ttu-id="a384f-146">Exemplo de VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="a384f-146">VersionOverrides example</span></span>

<span data-ttu-id="a384f-147">A seguir está um exemplo de uma típica `<VersionOverrides>` elemento, incluindo alguns elementos filhos que não são necessários, mas geralmente são usados.</span><span class="sxs-lookup"><span data-stu-id="a384f-147">The following is an example of a typical `<VersionOverrides>` element, including some child elements that are not required but are typically used.</span></span>

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

## <a name="implementing-multiple-versions"></a><span data-ttu-id="a384f-148">Implementar várias versões</span><span class="sxs-lookup"><span data-stu-id="a384f-148">Implementing multiple versions</span></span>

<span data-ttu-id="a384f-p107">Um manifesto pode implementar várias versões do elemento `VersionOverrides` que é compatível com várias versões do esquema VersionOverrides. Isso pode ser feito para fornecer suporte opcional a novos recursos em um esquema mais recente, sem deixar de fornecer suporte a clientes antigos que não têm suporte para os novos recursos.</span><span class="sxs-lookup"><span data-stu-id="a384f-p107">A manifest can implement multiple versions of the `VersionOverrides` element which support different versions of the VersionOverrides schema. This can be done to optionally support new features in a newer schema while still supporting older clients that do not support the new features.</span></span>

<span data-ttu-id="a384f-151">Para implementar várias versões, o elemento `VersionOverrides` da versão mais recente deve ser um filho do elemento `VersionOverrides` da versão anterior.</span><span class="sxs-lookup"><span data-stu-id="a384f-151">In order to implement multiple versions, the `VersionOverrides` element for the newer version must be a child of the `VersionOverrides` element for the older version.</span></span> <span data-ttu-id="a384f-152">O elemento filho `VersionOverrides` não herda os valores do elemento pai.</span><span class="sxs-lookup"><span data-stu-id="a384f-152">The child `VersionOverrides` element doesn't inherit any values from the parent.</span></span>

<span data-ttu-id="a384f-153">Para implementar o esquema do VersionOverrides v1.0 e do v1.1, o manifesto seria semelhante ao exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="a384f-153">To implement both the VersionOverrides v1.0 and v1.1 schema, the manifest would look similar to the following example:</span></span>

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
