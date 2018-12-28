---
title: Elemento AllFormFactors no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: de7fcdce48e175d15ca6268f24082e37b2085b05
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433275"
---
# <a name="allformfactors-element"></a><span data-ttu-id="1fcc6-102">Elemento AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="1fcc6-102">AllFormFactors element</span></span>

<span data-ttu-id="1fcc6-103">Especifica as configurações de um suplemento para todos os fatores forma.</span><span class="sxs-lookup"><span data-stu-id="1fcc6-103">Specifies the settings for an add-in for all form factors.</span></span> <span data-ttu-id="1fcc6-104">Atualmente, o único recurso que usa **AllFormFactors** são as funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="1fcc6-104">Currently, the only feature using AllFormFactors is custom functions.</span></span> <span data-ttu-id="1fcc6-105">**AllFormFactors** é um elemento obrigatório ao usar as funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="1fcc6-105">AllFormFactors is a required element when using custom functions.</span></span>

## <a name="child-elements"></a><span data-ttu-id="1fcc6-106">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="1fcc6-106">Child elements</span></span>

|  <span data-ttu-id="1fcc6-107">Elemento</span><span class="sxs-lookup"><span data-stu-id="1fcc6-107">Element</span></span> |  <span data-ttu-id="1fcc6-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="1fcc6-108">Required</span></span>  |  <span data-ttu-id="1fcc6-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="1fcc6-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="1fcc6-110">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="1fcc6-110">ExtensionPoint</span></span>](extensionpoint.md) |  <span data-ttu-id="1fcc6-111">Sim</span><span class="sxs-lookup"><span data-stu-id="1fcc6-111">Yes</span></span> |  <span data-ttu-id="1fcc6-112">Define onde um suplemento expõe a funcionalidade.</span><span class="sxs-lookup"><span data-stu-id="1fcc6-112">Defines where an add-in exposes functionality.</span></span> |

## <a name="allformfactors-example"></a><span data-ttu-id="1fcc6-113">Exemplo de AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="1fcc6-113">AllFormFactors example</span></span>

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
