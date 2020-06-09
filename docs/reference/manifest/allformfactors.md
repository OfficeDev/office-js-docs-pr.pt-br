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
