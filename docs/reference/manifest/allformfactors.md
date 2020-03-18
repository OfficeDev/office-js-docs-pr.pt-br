---
title: Elemento AllFormFactors no arquivo de manifesto
description: Especifica as configurações de um suplemento para todos os fatores forma.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: f1285f92b5eb89993e7fcfe79aab2325b86aca3d
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720712"
---
# <a name="allformfactors-element"></a><span data-ttu-id="931a2-103">Elemento AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="931a2-103">AllFormFactors element</span></span>

<span data-ttu-id="931a2-104">Especifica as configurações de um suplemento para todos os fatores forma.</span><span class="sxs-lookup"><span data-stu-id="931a2-104">Specifies the settings for an add-in for all form factors.</span></span> <span data-ttu-id="931a2-105">Atualmente, o único recurso que usa **AllFormFactors** são as funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="931a2-105">Currently, the only feature using **AllFormFactors** is custom functions.</span></span> <span data-ttu-id="931a2-106">**AllFormFactors** é um elemento obrigatório ao usar as funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="931a2-106">**AllFormFactors** is a required element when using custom functions.</span></span>

## <a name="child-elements"></a><span data-ttu-id="931a2-107">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="931a2-107">Child elements</span></span>

|  <span data-ttu-id="931a2-108">Elemento</span><span class="sxs-lookup"><span data-stu-id="931a2-108">Element</span></span> |  <span data-ttu-id="931a2-109">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="931a2-109">Required</span></span>  |  <span data-ttu-id="931a2-110">Descrição</span><span class="sxs-lookup"><span data-stu-id="931a2-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="931a2-111">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="931a2-111">ExtensionPoint</span></span>](extensionpoint.md) |  <span data-ttu-id="931a2-112">Sim</span><span class="sxs-lookup"><span data-stu-id="931a2-112">Yes</span></span> |  <span data-ttu-id="931a2-113">Define onde um suplemento expõe a funcionalidade.</span><span class="sxs-lookup"><span data-stu-id="931a2-113">Defines where an add-in exposes functionality.</span></span> |

## <a name="allformfactors-example"></a><span data-ttu-id="931a2-114">Exemplo de AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="931a2-114">AllFormFactors example</span></span>

```xml
<Hosts>
    <Host xsi:type="Workbook">
        <AllFormFactors>
            <ExtensionPoint xsi:type="CustomFunctions">
                    <!-- Information on this extension point -->
            </ExtensionPoint>
        </AllFormFactors>
    </Host>
</Hosts>
```
