---
title: Elemento DesktopFormFactor no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: dea632f7f8afa5d9b69f257798022e9e520e9394
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433737"
---
# <a name="desktopformfactor-element"></a>Elemento DesktopFormFactor

Especifica as configurações de um suplemento para o fator forma da área de trabalho. O fator de forma da área de trabalho inclui o Office para Windows, Office para Mac e Office Online. Ele contém todas as informações do suplemento para o fator forma da área de trabalho, exceto para o nó **Resources**.

Cada definição de DesktopFormFactor contém o elemento **FunctionFile** e um ou mais elementos **ExtensionPoint**. Para saber mais, confira [Elemento FunctionFile](functionfile.md) e [Elemento ExtensionPoint](extensionpoint.md).

## <a name="child-elements"></a>Elementos filho

| Elemento                               | Obrigatório | Descrição  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](extensionpoint.md)   | Sim      | Define onde um suplemento expõe a funcionalidade. |
| [FunctionFile](functionfile.md)       | Sim      | Uma URL para um arquivo que contém funções JavaScript.|
| [GetStarted](getstarted.md)           | Não       | Define o texto explicativo que aparece ao instalar o suplemento em hosts do Word, Excel ou PowerPoint. |
| [SupportsSharedFolders](supportssharedfolders.md) | Não | Define se o suplemento do Outlook está disponível em cenários de representante e é definido como *false* por padrão.<br><br>**Importante**: esse elemento só está disponível no conjunto de requisitos de visualização de suplementos do Outlook em comparação com o Exchange Online. Os suplementos que usam esse elemento não podem ser publicados no AppSource ou implantados por meio da implantação centralizada. |

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
        <!-- information on this extension point -->
      </ExtensionPoint>
      <!-- possibly more ExtensionPoint elements -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
