---
title: Elemento DesktopFormFactor no arquivo de manifesto
description: Especifica as configurações de um suplemento para o fator forma da área de trabalho.
ms.date: 09/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 52c9a029e3f43e9b7d5416455eb99ef3de4dae7a
ms.sourcegitcommit: 489befc41e543a4fb3c504fd9b3f61322134c1ef
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/06/2021
ms.locfileid: "60138727"
---
# <a name="desktopformfactor-element"></a>Elemento DesktopFormFactor

Especifica as configurações de um suplemento para o fator forma da área de trabalho. O fator de formulário da área de trabalho inclui Office na Web, Windows e Mac. Ele contém todas as informações de complemento para o fator de formulário da área de trabalho, exceto para o **nó Recursos.**

Cada definição desktopFormFactor contém o **elemento FunctionFile** e um ou mais **elementos ExtensionPoint.** Para saber mais, confira [Elemento FunctionFile](functionfile.md) e [Elemento ExtensionPoint](extensionpoint.md).

**Tipo de complemento:** Painel de tarefas, Email

**Válido somente nestes esquemas VersionOverrides:**

- Painel de tarefas 1.0
- Email 1.0
- Email 1.1

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

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
