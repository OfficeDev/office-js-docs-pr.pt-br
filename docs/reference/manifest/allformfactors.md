---
title: Elemento AllFormFactors no arquivo de manifesto
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 8059501f88f966b285398ac7cf243e6b0e4e44ea
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450734"
---
# <a name="allformfactors-element"></a><span data-ttu-id="efd8e-102">Elemento AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="efd8e-102">AllFormFactors element</span></span>

<span data-ttu-id="efd8e-103">Especifica as configurações de um suplemento para todos os fatores forma.</span><span class="sxs-lookup"><span data-stu-id="efd8e-103">Specifies the settings for an add-in for all form factors.</span></span> <span data-ttu-id="efd8e-104">Atualmente, o único recurso que usa **AllFormFactors** são as funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="efd8e-104">Currently, the only feature using **AllFormFactors** is custom functions.</span></span> <span data-ttu-id="efd8e-105">**AllFormFactors** é um elemento obrigatório ao usar as funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="efd8e-105">**AllFormFactors** is a required element when using custom functions.</span></span>

## <a name="child-elements"></a><span data-ttu-id="efd8e-106">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="efd8e-106">Child elements</span></span>

|  <span data-ttu-id="efd8e-107">Elemento</span><span class="sxs-lookup"><span data-stu-id="efd8e-107">Element</span></span> |  <span data-ttu-id="efd8e-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="efd8e-108">Required</span></span>  |  <span data-ttu-id="efd8e-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="efd8e-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="efd8e-110">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="efd8e-110">ExtensionPoint</span></span>](extensionpoint.md) |  <span data-ttu-id="efd8e-111">Sim</span><span class="sxs-lookup"><span data-stu-id="efd8e-111">Yes</span></span> |  <span data-ttu-id="efd8e-112">Define onde um suplemento expõe a funcionalidade.</span><span class="sxs-lookup"><span data-stu-id="efd8e-112">Defines where an add-in exposes functionality.</span></span> |

## <a name="allformfactors-example"></a><span data-ttu-id="efd8e-113">Exemplo de AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="efd8e-113">AllFormFactors example</span></span>

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
