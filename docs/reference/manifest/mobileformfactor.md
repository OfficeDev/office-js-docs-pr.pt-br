---
title: Elemento MobileFormFactor no arquivo de manifesto
description: O elemento MobileFormFactor especifica as configurações do fator de formulário móvel para um suplemento.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 954fff5d1e701ce53a6ad82fa276c048ca6d6f3a
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720586"
---
# <a name="mobileformfactor-element"></a>Elemento MobileFormFactor

Especifica as configurações de um suplemento para um fator forma móvel. Ele contém todas as informações do suplemento para o fator forma móvel, exceto para o nó **Resources**.

Cada definição de **MobileFormFactor** contém o elemento **functionfile** e um ou mais elementos **ExtensionPoint** . Para saber mais, confira [Elemento FunctionFile](functionfile.md) e [Elemento ExtensionPoint](extensionpoint.md).

O elemento **MobileFormFactor** é definido no esquema VersionOverrides 1.1. O elemento [VersionOverrides](versionoverrides.md) incluído deve ter um valor de atributo `xsi:type` de `VersionOverridesV1_1`.

## <a name="child-elements"></a>Elementos filho

| Elemento                               | Obrigatório | Descrição  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](extensionpoint.md) | Sim      | Define onde um suplemento expõe a funcionalidade. |
| [FunctionFile](functionfile.md)     | Sim      | Uma URL para um arquivo que contém funções JavaScript.|

## <a name="mobileformfactor-example"></a>Exemplo de MobileFormFactor

```xml
...
<Hosts>
  <Host xsi:type="MailHost">
    ...
    <MobileFormFactor>
      <FunctionFile resid="residUILessFunctionFileUrl" />
      <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint> 
      <!-- possibly more ExtensionPoint elements -->
    </MobileFormFactor>
  </Host>
</Hosts>
...
```
