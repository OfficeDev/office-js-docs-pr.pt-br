---
title: Elemento DesktopFormFactor no arquivo de manifesto
description: Especifica as configurações de um suplemento para o fator forma da área de trabalho.
ms.date: 09/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: 3f15840a7b6716cd8acabe9e061effa566d48930
ms.sourcegitcommit: a854a2fd2ad9f379a3ef712f307e0b1bb9b5b00d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/22/2021
ms.locfileid: "59474326"
---
# <a name="desktopformfactor-element"></a>Elemento DesktopFormFactor

Especifica as configurações de um suplemento para o fator forma da área de trabalho. O fator de formulário da área de trabalho inclui Office na Web, Windows e Mac. Ele contém todas as informações de complemento para o fator de formulário da área de trabalho, exceto para o **nó Recursos.**

Cada definição desktopFormFactor contém o **elemento FunctionFile** e um ou mais **elementos ExtensionPoint.** Para saber mais, confira [Elemento FunctionFile](functionfile.md) e [Elemento ExtensionPoint](extensionpoint.md).

## <a name="child-elements"></a>Elementos filho

| Elemento                               | Obrigatório | Descrição  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](extensionpoint.md)   | Sim      | Define onde um suplemento expõe a funcionalidade. |
| [FunctionFile](functionfile.md)       | Sim      | Uma URL para um arquivo que contém funções JavaScript.|
| [GetStarted](getstarted.md)           | Não       | Define o texto explicante que aparece ao instalar o complemento no Word, Excel ou PowerPoint. Se omitido, o callout usará os valores dos elementos [DisplayName](displayname.md) e [Description.](description.md) |
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
