---
title: Elemento AllFormFactors no arquivo de manifesto
description: Especifica as configurações de um suplemento para todos os fatores forma.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 9dac322312c1dfd60f6deb4296413e12b55a6a49
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608793"
---
# <a name="allformfactors-element"></a><span data-ttu-id="b6976-103">Elemento AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="b6976-103">AllFormFactors element</span></span>

<span data-ttu-id="b6976-104">Especifica as configurações de um suplemento para todos os fatores forma.</span><span class="sxs-lookup"><span data-stu-id="b6976-104">Specifies the settings for an add-in for all form factors.</span></span> <span data-ttu-id="b6976-105">Atualmente, o único recurso que usa **AllFormFactors** são as funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="b6976-105">Currently, the only feature using **AllFormFactors** is custom functions.</span></span> <span data-ttu-id="b6976-106">**AllFormFactors** é um elemento obrigatório ao usar as funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="b6976-106">**AllFormFactors** is a required element when using custom functions.</span></span>

## <a name="child-elements"></a><span data-ttu-id="b6976-107">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="b6976-107">Child elements</span></span>

|  <span data-ttu-id="b6976-108">Elemento</span><span class="sxs-lookup"><span data-stu-id="b6976-108">Element</span></span> |  <span data-ttu-id="b6976-109">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="b6976-109">Required</span></span>  |  <span data-ttu-id="b6976-110">Descrição</span><span class="sxs-lookup"><span data-stu-id="b6976-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="b6976-111">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="b6976-111">ExtensionPoint</span></span>](extensionpoint.md) |  <span data-ttu-id="b6976-112">Sim</span><span class="sxs-lookup"><span data-stu-id="b6976-112">Yes</span></span> |  <span data-ttu-id="b6976-113">Define onde um suplemento expõe a funcionalidade.</span><span class="sxs-lookup"><span data-stu-id="b6976-113">Defines where an add-in exposes functionality.</span></span> |

## <a name="allformfactors-example"></a><span data-ttu-id="b6976-114">Exemplo de AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="b6976-114">AllFormFactors example</span></span>

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
