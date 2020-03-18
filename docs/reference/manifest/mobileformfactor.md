---
title: Elemento MobileFormFactor no arquivo de manifesto
description: O elemento MobileFormFactor especifica as configurações do fator de formulário móvel para um suplemento.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 954fff5d1e701ce53a6ad82fa276c048ca6d6f3a
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720586"
---
# <a name="mobileformfactor-element"></a><span data-ttu-id="51a20-103">Elemento MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="51a20-103">MobileFormFactor element</span></span>

<span data-ttu-id="51a20-p101">Especifica as configurações de um suplemento para um fator forma móvel. Ele contém todas as informações do suplemento para o fator forma móvel, exceto para o nó **Resources**.</span><span class="sxs-lookup"><span data-stu-id="51a20-p101">Specifies the settings for an add-in for the mobile form factor. It contains all the add-in information for the mobile form factor except for the **Resources** node.</span></span>

<span data-ttu-id="51a20-106">Cada definição de **MobileFormFactor** contém o elemento **functionfile** e um ou mais elementos **ExtensionPoint** .</span><span class="sxs-lookup"><span data-stu-id="51a20-106">Each **MobileFormFactor** definition contains the **FunctionFile** element and one or more **ExtensionPoint** elements.</span></span> <span data-ttu-id="51a20-107">Para saber mais, confira [Elemento FunctionFile](functionfile.md) e [Elemento ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="51a20-107">For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

<span data-ttu-id="51a20-p103">O elemento **MobileFormFactor** é definido no esquema VersionOverrides 1.1. O elemento [VersionOverrides](versionoverrides.md) incluído deve ter um valor de atributo `xsi:type` de `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="51a20-p103">The **MobileFormFactor** element is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

## <a name="child-elements"></a><span data-ttu-id="51a20-110">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="51a20-110">Child elements</span></span>

| <span data-ttu-id="51a20-111">Elemento</span><span class="sxs-lookup"><span data-stu-id="51a20-111">Element</span></span>                               | <span data-ttu-id="51a20-112">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="51a20-112">Required</span></span> | <span data-ttu-id="51a20-113">Descrição</span><span class="sxs-lookup"><span data-stu-id="51a20-113">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="51a20-114">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="51a20-114">ExtensionPoint</span></span>](extensionpoint.md) | <span data-ttu-id="51a20-115">Sim</span><span class="sxs-lookup"><span data-stu-id="51a20-115">Yes</span></span>      | <span data-ttu-id="51a20-116">Define onde um suplemento expõe a funcionalidade.</span><span class="sxs-lookup"><span data-stu-id="51a20-116">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="51a20-117">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="51a20-117">FunctionFile</span></span>](functionfile.md)     | <span data-ttu-id="51a20-118">Sim</span><span class="sxs-lookup"><span data-stu-id="51a20-118">Yes</span></span>      | <span data-ttu-id="51a20-119">Uma URL para um arquivo que contém funções JavaScript.</span><span class="sxs-lookup"><span data-stu-id="51a20-119">A URL to a file that contains JavaScript functions.</span></span>|

## <a name="mobileformfactor-example"></a><span data-ttu-id="51a20-120">Exemplo de MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="51a20-120">MobileFormFactor example</span></span>

```xml
...
<Hosts>
  <Host xsi:type="MailHost">
    ...
    <MobileFormFactor>
      <FunctionFile resid="residUILessFunctionFileUrl" />
      <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint> 
      <!-- possibly more ExtensionPoint elements -->
    </MobileFormFactor>
  </Host>
</Hosts>
...
```
