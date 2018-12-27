---
title: Elemento MobileFormFactor no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: f0a68c7127f7872207a58ed252def7a2977c33ed
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433688"
---
# <a name="mobileformfactor-element"></a><span data-ttu-id="500b3-102">Elemento MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="500b3-102">MobileFormFactor element</span></span>

<span data-ttu-id="500b3-p101">Especifica as configurações de um suplemento para um fator forma móvel. Ele contém todas as informações do suplemento para o fator forma móvel, exceto para o nó **Resources**.</span><span class="sxs-lookup"><span data-stu-id="500b3-p101">Specifies the settings for an add-in for the mobile form factor. It contains all the add-in information for the mobile form factor except for the **Resources** node.</span></span>

<span data-ttu-id="500b3-p102">Cada definição de **MobileFormFactor** contém o elemento **FunctionFile** e um ou mais elementos **ExtensionPoint**. Para saber mais, confira [Elemento FunctionFile](functionfile.md) e [Elemento ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="500b3-p102">Each **MobileFormFactor** definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

<span data-ttu-id="500b3-p103">O elemento **MobileFormFactor** é definido no esquema VersionOverrides 1.1. O elemento [VersionOverrides](versionoverrides.md) incluído deve ter um valor de atributo `xsi:type` de `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="500b3-p103">The **MobileFormFactor** element is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

## <a name="child-elements"></a><span data-ttu-id="500b3-109">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="500b3-109">Child elements</span></span>

| <span data-ttu-id="500b3-110">Elemento</span><span class="sxs-lookup"><span data-stu-id="500b3-110">Element</span></span>                               | <span data-ttu-id="500b3-111">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="500b3-111">Required</span></span> | <span data-ttu-id="500b3-112">Descrição</span><span class="sxs-lookup"><span data-stu-id="500b3-112">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="500b3-113">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="500b3-113">ExtensionPoint</span></span>](extensionpoint.md) | <span data-ttu-id="500b3-114">Sim</span><span class="sxs-lookup"><span data-stu-id="500b3-114">Yes</span></span>      | <span data-ttu-id="500b3-115">Define onde um suplemento expõe a funcionalidade.</span><span class="sxs-lookup"><span data-stu-id="500b3-115">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="500b3-116">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="500b3-116">FunctionFile</span></span>](functionfile.md)     | <span data-ttu-id="500b3-117">Sim</span><span class="sxs-lookup"><span data-stu-id="500b3-117">Yes</span></span>      | <span data-ttu-id="500b3-118">Uma URL para um arquivo que contém funções JavaScript.</span><span class="sxs-lookup"><span data-stu-id="500b3-118">A URL to a file that contains JavaScript functions.</span></span>|

## <a name="mobileformfactor-example"></a><span data-ttu-id="500b3-119">Exemplo de MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="500b3-119">MobileFormFactor example</span></span>

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
