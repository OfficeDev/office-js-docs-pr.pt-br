---
title: Elemento DesktopFormFactor no arquivo de manifesto
description: Especifica as configurações de um suplemento para o fator forma da área de trabalho.
ms.date: 06/15/2021
ms.localizationpriority: medium
ms.openlocfilehash: f89dff5626867258c8df93d5f047e3d08103e71b
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59149057"
---
# <a name="desktopformfactor-element"></a>Elemento DesktopFormFactor

Especifica as configurações de um suplemento para o fator forma da área de trabalho. O fator de formulário da área de trabalho inclui Office na Web, Windows e Mac. Ele contém todas as informações de complemento para o fator de formulário da área de trabalho, exceto para o **nó Recursos.**

Cada definição desktopFormFactor contém o **elemento FunctionFile** e um ou mais **elementos ExtensionPoint.** Para saber mais, confira [Elemento FunctionFile](functionfile.md) e [Elemento ExtensionPoint](extensionpoint.md).

## <a name="child-elements"></a>Elementos filho

| Elemento                               | Obrigatório | Descrição  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](extensionpoint.md)   | Sim      | Define onde um suplemento expõe a funcionalidade. |
| [FunctionFile](functionfile.md)       | Sim      | Uma URL para um arquivo que contém funções JavaScript.|
| [GetStarted](getstarted.md)           | Não       | Define o texto explicante que aparece ao instalar o complemento no Word, Excel ou PowerPoint. |
| [SupportsSharedFolders](supportssharedfolders.md) | Não | Define se o Outlook está disponível em cenários de caixa de correio compartilhada (agora em visualização) e pastas compartilhadas (ou seja, acesso de representante). Definir como *false* por padrão. |

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
