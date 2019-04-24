---
title: Elemento DesktopFormFactor no arquivo de manifesto
description: ''
ms.date: 03/01/2019
localization_priority: Normal
ms.openlocfilehash: cddf76af01ec9f3016b28a3f7692aa6dfeb9bd60
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450552"
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
| [SupportsSharedFolders](supportssharedfolders.md) | Não | Define se o suplemento do Outlook está disponível em cenários de representante e é definido como *false* por padrão.<br><br>**Importante**: como o acesso de representante para suplementos do Outlook está atualmente em versão prévia, os suplementos que usam `SupportSharedFolders` o elemento não podem ser publicados no AppSource ou implantados por meio da implantação centralizada. |

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
