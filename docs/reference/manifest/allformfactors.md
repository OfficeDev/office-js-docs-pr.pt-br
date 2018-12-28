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
