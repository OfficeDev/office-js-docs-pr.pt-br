---
title: Elemento AllFormFactors no arquivo de manifesto
description: Especifica as configurações de um suplemento para todos os fatores forma.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 674fbe9defa961cb0eef1103cf2dedea0983ffabadc665b172d1f3b15292e987
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57088534"
---
# <a name="allformfactors-element"></a>Elemento AllFormFactors

Especifica as configurações de um suplemento para todos os fatores forma. Atualmente, o único recurso que usa **AllFormFactors** são as funções personalizadas. **AllFormFactors** é um elemento obrigatório ao usar as funções personalizadas.

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [ExtensionPoint](extensionpoint.md) |  Sim |  Define onde um suplemento expõe a funcionalidade. |

## <a name="allformfactors-example"></a>Exemplo de AllFormFactors

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
