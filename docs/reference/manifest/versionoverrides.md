---
title: Elemento VersionOverrides no arquivo de manifesto
description: Documentação de referência do elemento VersionOverrides para arquivos de manifesto de suplementos do Office (XML).
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: 772eaa416909d24f8035ed3e1445d1e4f06a244e
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771302"
---
# <a name="versionoverrides-element"></a><span data-ttu-id="b018e-103">Elemento VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="b018e-103">VersionOverrides element</span></span>

<span data-ttu-id="b018e-p101">O elemento raiz que contém informações para os comandos de suplemento implementados pelo suplemento. **VersionOverrides** é um elemento filho do elemento [OfficeApp](officeapp.md) no manifesto. Ele recebe suporte no esquema de manifesto v1.1 e posterior, mas é definido no esquema VersionOverrides v1.0 ou v1.1.</span><span class="sxs-lookup"><span data-stu-id="b018e-p101">The root element that contains information for the add-in commands implemented by the add-in. **VersionOverrides** is a child element of the [OfficeApp](officeapp.md) element in the manifest. This element is supported in manifest schema v1.1 and later but is defined in the VersionOverrides v1.0 or v1.1 schema.</span></span>

## <a name="attributes"></a><span data-ttu-id="b018e-107">Atributos</span><span class="sxs-lookup"><span data-stu-id="b018e-107">Attributes</span></span>

|  <span data-ttu-id="b018e-108">Atributo</span><span class="sxs-lookup"><span data-stu-id="b018e-108">Attribute</span></span>  |  <span data-ttu-id="b018e-109">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="b018e-109">Required</span></span>  |  <span data-ttu-id="b018e-110">Descrição</span><span class="sxs-lookup"><span data-stu-id="b018e-110">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="b018e-111">**xmlns**</span><span class="sxs-lookup"><span data-stu-id="b018e-111">**xmlns**</span></span>       |  <span data-ttu-id="b018e-112">Sim</span><span class="sxs-lookup"><span data-stu-id="b018e-112">Yes</span></span>  |  <span data-ttu-id="b018e-113">O namespace do esquema VersionOverrides.</span><span class="sxs-lookup"><span data-stu-id="b018e-113">The VersionOverrides schema namespace.</span></span> <span data-ttu-id="b018e-114">Os valores permitidos variam de acordo com o `<VersionOverrides>` valor **xsi: Type** do elemento e o valor **xsi: Type** do elemento pai `<OfficeApp>` .</span><span class="sxs-lookup"><span data-stu-id="b018e-114">The allowed values vary depending on  this `<VersionOverrides>` element's **xsi:type** value and the **xsi:type** value of the parent `<OfficeApp>` element.</span></span> <span data-ttu-id="b018e-115">Consulte [namespace valores](#namespace-values) a seguir.</span><span class="sxs-lookup"><span data-stu-id="b018e-115">See [Namespace values](#namespace-values) below.</span></span>|
|  <span data-ttu-id="b018e-116">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="b018e-116">**xsi:type**</span></span>  |  <span data-ttu-id="b018e-117">Sim</span><span class="sxs-lookup"><span data-stu-id="b018e-117">Yes</span></span>  | <span data-ttu-id="b018e-p103">A versão do esquema. Nesse momento, os únicos valores válidos são `VersionOverridesV1_0` e `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="b018e-p103">The schema version. At this time, the only valid values are `VersionOverridesV1_0` and `VersionOverridesV1_1`.</span></span> |

### <a name="namespace-values"></a><span data-ttu-id="b018e-120">Valores do namespace</span><span class="sxs-lookup"><span data-stu-id="b018e-120">Namespace values</span></span>

<span data-ttu-id="b018e-121">A seguir, a lista o valor necessário do valor **xmlns** , dependendo do valor **xsi: Type** do elemento pai `<OfficeApp>` .</span><span class="sxs-lookup"><span data-stu-id="b018e-121">The following lists the required value of the **xmlns** value depending on the **xsi:type** value of the parent `<OfficeApp>` element.</span></span>

- <span data-ttu-id="b018e-122">**TaskPaneApp** oferece suporte somente à versão 1,0 do VersionOverrides e o **xmlns** deve ser `http://schemas.microsoft.com/office/taskpaneappversionoverrides` .</span><span class="sxs-lookup"><span data-stu-id="b018e-122">**TaskPaneApp** supports only version 1.0 of VersionOverrides, and the **xmlns** should be `http://schemas.microsoft.com/office/taskpaneappversionoverrides`.</span></span>
- <span data-ttu-id="b018e-123">**ContentApp** oferece suporte somente à versão 1,0 do VersionOverrides e o **xmlns** deve ser `http://schemas.microsoft.com/office/contentappversionoverrides` .</span><span class="sxs-lookup"><span data-stu-id="b018e-123">**ContentApp** supports only version 1.0 of VersionOverrides, and the **xmlns** should be `http://schemas.microsoft.com/office/contentappversionoverrides`.</span></span>
- <span data-ttu-id="b018e-124">O **MailApp** suporta as versões 1,0 e 1,1 do VersionOverrides, portanto, o valor de **xmlns** varia de acordo com o `<VersionOverrides>` valor **xsi: Type** do elemento:</span><span class="sxs-lookup"><span data-stu-id="b018e-124">**MailApp** supports versions 1.0 and 1.1 of VersionOverrides, so the value of **xmlns** varies depending on this `<VersionOverrides>` element's **xsi:type** value:</span></span>
    - <span data-ttu-id="b018e-125">Quando **xsi: Type** é `VersionOverridesV1_0` , o **xmlns** deve ser `http://schemas.microsoft.com/office/mailappversionoverrides` .</span><span class="sxs-lookup"><span data-stu-id="b018e-125">When **xsi:type** is `VersionOverridesV1_0`, **xmlns** must be `http://schemas.microsoft.com/office/mailappversionoverrides`.</span></span>
    - <span data-ttu-id="b018e-126">Quando **xsi: Type** é `VersionOverridesV1_1` , o **xmlns** deve ser `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` .</span><span class="sxs-lookup"><span data-stu-id="b018e-126">When **xsi:type** is `VersionOverridesV1_1`, **xmlns** must be `http://schemas.microsoft.com/office/mailappversionoverrides/1.1`.</span></span>

> [!NOTE]
> <span data-ttu-id="b018e-127">Atualmente, somente o Outlook 2016 ou posterior oferece suporte ao esquema do VersionOverrides v 1.1 e ao `VersionOverridesV1_1` tipo.</span><span class="sxs-lookup"><span data-stu-id="b018e-127">Currently only Outlook 2016 or later supports the VersionOverrides v1.1 schema and the `VersionOverridesV1_1` type.</span></span>

## <a name="child-elements"></a><span data-ttu-id="b018e-128">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="b018e-128">Child elements</span></span>

|  <span data-ttu-id="b018e-129">Elemento</span><span class="sxs-lookup"><span data-stu-id="b018e-129">Element</span></span> |  <span data-ttu-id="b018e-130">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="b018e-130">Required</span></span>  |  <span data-ttu-id="b018e-131">Descrição</span><span class="sxs-lookup"><span data-stu-id="b018e-131">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="b018e-132">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="b018e-132">**Description**</span></span>    |  <span data-ttu-id="b018e-133">Não</span><span class="sxs-lookup"><span data-stu-id="b018e-133">No</span></span>   |  <span data-ttu-id="b018e-134">Descreve o suplemento.</span><span class="sxs-lookup"><span data-stu-id="b018e-134">Describes the add-in.</span></span> <span data-ttu-id="b018e-135">Isso substitui o elemento `Description` em qualquer parte pai do manifesto.</span><span class="sxs-lookup"><span data-stu-id="b018e-135">This overrides the `Description` element in any parent portion of the manifest.</span></span> <span data-ttu-id="b018e-136">O texto da descrição está contido em um elemento filho do elemento **LongString**, contido no elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="b018e-136">The text of the description is contained in a child element of the **LongString** element contained in the [Resources](resources.md) element.</span></span> <span data-ttu-id="b018e-137">O `resid` atributo do elemento **Description** não pode ter mais de 32 caracteres e é definido como o valor do `id` atributo do `String` elemento que contém o texto.</span><span class="sxs-lookup"><span data-stu-id="b018e-137">The `resid` attribute of the **Description** element can be no more than 32 characters and is set to the value of the `id` attribute of the `String` element that contains the text.</span></span>|
|  <span data-ttu-id="b018e-138">**Requisitos**</span><span class="sxs-lookup"><span data-stu-id="b018e-138">**Requirements**</span></span>  |  <span data-ttu-id="b018e-139">Não</span><span class="sxs-lookup"><span data-stu-id="b018e-139">No</span></span>   |  <span data-ttu-id="b018e-p105">Especifica o conjunto de requisitos mínimos e a versão do Office.js exigida pelo suplemento. Isso substitui o elemento `Requirements` na parte pai do manifesto.</span><span class="sxs-lookup"><span data-stu-id="b018e-p105">Specifies the minimum requirement set and version of Office.js that the add-in requires. This overrides the  `Requirements` element in the parent portion of the manifest.</span></span>|
|  [<span data-ttu-id="b018e-142">Hosts</span><span class="sxs-lookup"><span data-stu-id="b018e-142">Hosts</span></span>](hosts.md)                |  <span data-ttu-id="b018e-143">Sim</span><span class="sxs-lookup"><span data-stu-id="b018e-143">Yes</span></span>  |  <span data-ttu-id="b018e-144">Especifica uma coleção de aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="b018e-144">Specifies a collection of Office applications.</span></span> <span data-ttu-id="b018e-145">O elemento hosts filho substitui o elemento hosts na parte pai do manifesto.</span><span class="sxs-lookup"><span data-stu-id="b018e-145">The child Hosts element overrides the Hosts element in the parent portion of the manifest.</span></span>  |
|  [<span data-ttu-id="b018e-146">Resources</span><span class="sxs-lookup"><span data-stu-id="b018e-146">Resources</span></span>](resources.md)    |  <span data-ttu-id="b018e-147">Sim</span><span class="sxs-lookup"><span data-stu-id="b018e-147">Yes</span></span>  | <span data-ttu-id="b018e-148">Define um conjunto de recursos (cadeias de caracteres, URLs e imagens) consultado por outros elementos do manifesto.</span><span class="sxs-lookup"><span data-stu-id="b018e-148">Defines a collection of resources (strings, URLs, and images) that other manifest elements reference.</span></span>|
|  [<span data-ttu-id="b018e-149">EquivalentAddins</span><span class="sxs-lookup"><span data-stu-id="b018e-149">EquivalentAddins</span></span>](equivalentaddins.md)    |  <span data-ttu-id="b018e-150">Não</span><span class="sxs-lookup"><span data-stu-id="b018e-150">No</span></span>  | <span data-ttu-id="b018e-151">Especifica os suplementos nativos (COM/XLL) equivalentes ao suplemento Web.</span><span class="sxs-lookup"><span data-stu-id="b018e-151">Specifies the native (COM/XLL) add-ins that are equivalent to the web add-in.</span></span> <span data-ttu-id="b018e-152">O suplemento Web não será ativado se um suplemento nativo equivalente estiver instalado.</span><span class="sxs-lookup"><span data-stu-id="b018e-152">The web add-in is not activated if an equivalent native add-in is installed.</span></span>|
|  <span data-ttu-id="b018e-153">**VersionOverrides**</span><span class="sxs-lookup"><span data-stu-id="b018e-153">**VersionOverrides**</span></span>    |  <span data-ttu-id="b018e-154">Não</span><span class="sxs-lookup"><span data-stu-id="b018e-154">No</span></span>  | <span data-ttu-id="b018e-p108">Define comandos de suplemento em uma versão mais recente do esquema. Para saber mais, confira o tópico [Implementar várias versões](#implementing-multiple-versions).</span><span class="sxs-lookup"><span data-stu-id="b018e-p108">Defines add-in commands under a newer schema version. See [Implementing multiple versions](#implementing-multiple-versions) for details.</span></span> |
|  [<span data-ttu-id="b018e-157">WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="b018e-157">WebApplicationInfo</span></span>](webapplicationinfo.md)    |  <span data-ttu-id="b018e-158">Não</span><span class="sxs-lookup"><span data-stu-id="b018e-158">No</span></span>  | <span data-ttu-id="b018e-159">Especifica detalhes sobre o registro do suplemento com emissores de token seguros, como o Azure Active Directory V 2.0.</span><span class="sxs-lookup"><span data-stu-id="b018e-159">Specifies details about the add-in's registration with secure token issuers, such as Azure Active Directory V2.0.</span></span> |
|  [<span data-ttu-id="b018e-160">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="b018e-160">ExtendedPermissions</span></span>](extendedpermissions.md) |  <span data-ttu-id="b018e-161">Não</span><span class="sxs-lookup"><span data-stu-id="b018e-161">No</span></span>  |  <span data-ttu-id="b018e-162">Especifica uma coleção de permissões estendidas.</span><span class="sxs-lookup"><span data-stu-id="b018e-162">Specifies a collection of extended permissions.</span></span><br><br><span data-ttu-id="b018e-163">**Importante**: como a API [Office. Body. appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendonsendasync-data--options--callback-) está atualmente em versão prévia, os suplementos que usam o `ExtendedPermissions` elemento não podem ser publicados no AppSource ou implantados por meio da implantação centralizada.</span><span class="sxs-lookup"><span data-stu-id="b018e-163">**Important**: Because the [Office.Body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendonsendasync-data--options--callback-) API is currently in preview, add-ins that use the `ExtendedPermissions` element can't be published to AppSource or deployed via centralized deployment.</span></span> |

### <a name="versionoverrides-example"></a><span data-ttu-id="b018e-164">Exemplo de VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="b018e-164">VersionOverrides example</span></span>

<span data-ttu-id="b018e-165">Veja a seguir um exemplo de um `<VersionOverrides>` elemento típico, incluindo alguns elementos filhos que não são necessários, mas que são normalmente usados.</span><span class="sxs-lookup"><span data-stu-id="b018e-165">The following is an example of a typical `<VersionOverrides>` element, including some child elements that are not required but are typically used.</span></span>

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

## <a name="implementing-multiple-versions"></a><span data-ttu-id="b018e-166">Implementar várias versões</span><span class="sxs-lookup"><span data-stu-id="b018e-166">Implementing multiple versions</span></span>

<span data-ttu-id="b018e-p109">Um manifesto pode implementar várias versões do elemento `VersionOverrides` que é compatível com várias versões do esquema VersionOverrides. Isso pode ser feito para fornecer suporte opcional a novos recursos em um esquema mais recente, sem deixar de fornecer suporte a clientes antigos que não têm suporte para os novos recursos.</span><span class="sxs-lookup"><span data-stu-id="b018e-p109">A manifest can implement multiple versions of the `VersionOverrides` element which support different versions of the VersionOverrides schema. This can be done to optionally support new features in a newer schema while still supporting older clients that do not support the new features.</span></span>

<span data-ttu-id="b018e-169">Para implementar várias versões, o elemento `VersionOverrides` da versão mais recente deve ser um filho do elemento `VersionOverrides` da versão anterior.</span><span class="sxs-lookup"><span data-stu-id="b018e-169">In order to implement multiple versions, the `VersionOverrides` element for the newer version must be a child of the `VersionOverrides` element for the older version.</span></span> <span data-ttu-id="b018e-170">O elemento filho `VersionOverrides` não herda os valores do elemento pai.</span><span class="sxs-lookup"><span data-stu-id="b018e-170">The child `VersionOverrides` element doesn't inherit any values from the parent.</span></span>

<span data-ttu-id="b018e-171">Para implementar o esquema do VersionOverrides v1.0 e do v1.1, o manifesto seria semelhante ao exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="b018e-171">To implement both the VersionOverrides v1.0 and v1.1 schema, the manifest would look similar to the following example:</span></span>

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
