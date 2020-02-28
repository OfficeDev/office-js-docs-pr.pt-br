---
title: Elemento MobileFormFactor no arquivo de manifesto
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 34106011cb855b6ac7c6d0fc21c16fd13e52b281
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324838"
---
# <a name="mobileformfactor-element"></a><span data-ttu-id="ae950-102">Elemento MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="ae950-102">MobileFormFactor element</span></span>

<span data-ttu-id="ae950-p101">Especifica as configurações de um suplemento para um fator forma móvel. Ele contém todas as informações do suplemento para o fator forma móvel, exceto para o nó **Resources**.</span><span class="sxs-lookup"><span data-stu-id="ae950-p101">Specifies the settings for an add-in for the mobile form factor. It contains all the add-in information for the mobile form factor except for the **Resources** node.</span></span>

<span data-ttu-id="ae950-105">Cada definição de **MobileFormFactor** contém o elemento **functionfile** e um ou mais elementos **ExtensionPoint** .</span><span class="sxs-lookup"><span data-stu-id="ae950-105">Each **MobileFormFactor** definition contains the **FunctionFile** element and one or more **ExtensionPoint** elements.</span></span> <span data-ttu-id="ae950-106">Para saber mais, confira [Elemento FunctionFile](functionfile.md) e [Elemento ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="ae950-106">For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

<span data-ttu-id="ae950-p103">O elemento **MobileFormFactor** é definido no esquema VersionOverrides 1.1. O elemento [VersionOverrides](versionoverrides.md) incluído deve ter um valor de atributo `xsi:type` de `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="ae950-p103">The **MobileFormFactor** element is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

## <a name="child-elements"></a><span data-ttu-id="ae950-109">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="ae950-109">Child elements</span></span>

| <span data-ttu-id="ae950-110">Elemento</span><span class="sxs-lookup"><span data-stu-id="ae950-110">Element</span></span>                               | <span data-ttu-id="ae950-111">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="ae950-111">Required</span></span> | <span data-ttu-id="ae950-112">Descrição</span><span class="sxs-lookup"><span data-stu-id="ae950-112">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="ae950-113">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="ae950-113">ExtensionPoint</span></span>](extensionpoint.md) | <span data-ttu-id="ae950-114">Sim</span><span class="sxs-lookup"><span data-stu-id="ae950-114">Yes</span></span>      | <span data-ttu-id="ae950-115">Define onde um suplemento expõe a funcionalidade.</span><span class="sxs-lookup"><span data-stu-id="ae950-115">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="ae950-116">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="ae950-116">FunctionFile</span></span>](functionfile.md)     | <span data-ttu-id="ae950-117">Sim</span><span class="sxs-lookup"><span data-stu-id="ae950-117">Yes</span></span>      | <span data-ttu-id="ae950-118">Uma URL para um arquivo que contém funções JavaScript.</span><span class="sxs-lookup"><span data-stu-id="ae950-118">A URL to a file that contains JavaScript functions.</span></span>|

## <a name="mobileformfactor-example"></a><span data-ttu-id="ae950-119">Exemplo de MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="ae950-119">MobileFormFactor example</span></span>

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
