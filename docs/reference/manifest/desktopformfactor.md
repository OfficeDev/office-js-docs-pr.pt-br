---
title: Elemento DesktopFormFactor no arquivo de manifesto
description: Especifica as configurações de um suplemento para o fator forma da área de trabalho.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: bfea6900e6b07d8dc7ad5b5256703d873242d88c
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718360"
---
# <a name="desktopformfactor-element"></a>Elemento DesktopFormFactor

Especifica as configurações de um suplemento para o fator forma da área de trabalho. O fator de forma da área de trabalho inclui o Office na Web, Windows e Mac. Ele contém todas as informações do suplemento para o fator de forma da área de trabalho, exceto para o nó de **recursos** .

Cada definição de DesktopFormFactor contém o elemento **functionfile** e um ou mais elementos **ExtensionPoint** . Para saber mais, confira [Elemento FunctionFile](functionfile.md) e [Elemento ExtensionPoint](extensionpoint.md).

## <a name="child-elements"></a>Elementos filho

| Elemento                               | Obrigatório | Descrição  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](extensionpoint.md)   | Sim      | Define onde um suplemento expõe a funcionalidade. |
| [FunctionFile](functionfile.md)       | Sim      | Uma URL para um arquivo que contém funções JavaScript.|
| [GetStarted](getstarted.md)           | Não       | Define o texto explicativo que aparece ao instalar o suplemento em hosts do Word, Excel ou PowerPoint. |
| [SupportsSharedFolders](supportssharedfolders.md) | Não | Define se o suplemento do Outlook está disponível em cenários de representante e é definido como *false* por padrão. |

## <a name="desktopformfactor-example"></a>Exemplo de DesktopFormFactor

```xml
...
<Hosts>
  <Host xsi:type="Presentation">
    <DesktopFormFactor>
      <FunctionFile resid="residDesktopFuncUrl" />
      <GetStarted>
        <!-- GetStarted callout -->
      </GetStarted>
      <ExtensionPoint xsi:type="PrimaryCommandSurface">
        <!-- Information on this extension point. -->
      </ExtensionPoint>
      <!-- Possibly more ExtensionPoint elements. -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
